using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Blocks;
using DocxTemplater.Formatter;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocxTemplater.Extensions;
using DocxTemplater.ImageBase;

namespace DocxTemplater
{
    public abstract class TemplateProcessor
    {
        internal ITemplateProcessingContextAccess Context { get; }

        private protected TemplateProcessor(ITemplateProcessingContextAccess context)
        {
            Context = context;
        }

        public IReadOnlyDictionary<string, object> Models => Context.ModelLookup.Models;


        protected void ProcessNode(OpenXmlCompositeElement rootElement)
        {
            var loops = BuildBlockTree(rootElement);
            RenderNode(rootElement, loops);
        }

        /// <summary>
        /// Renders the model into a node whose block tree has already been built by
        /// <see cref="BuildBlockTree"/>. This is the model-dependent half of the pipeline
        /// (variable replacement, extension rendering, loop/condition expansion, cleanup).
        /// Splitting it out lets <see cref="DocxTemplate.GetTemplateSchema"/> build the tree once
        /// and have <see cref="DocxTemplate.Process"/> reuse it instead of rebuilding (which would
        /// fail on the already-mutated tree).
        /// </summary>
        private protected void RenderNode(OpenXmlCompositeElement rootElement, IReadOnlyCollection<ContentBlock> loops)
        {
#if DEBUG
            Console.WriteLine("----------- After Loops --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
            Context.VariableReplacer.ReplaceVariables(rootElement, Context);
            foreach (var extensions in Context.Extensions)
            {
                extensions.ReplaceVariables(Context, rootElement, [.. rootElement]);
            }
            foreach (var loop in loops)
            {
                loop.Expand(Context.ModelLookup, rootElement);
            }

            Cleanup(rootElement, removeEmptyElements: true);
#if DEBUG
            Console.WriteLine("----------- Completed --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
        }

        private void PreProcess(OpenXmlCompositeElement content)
        {
            // remove spell check 'ProofError' elements
            content.Descendants<ProofError>().ToList().ForEach(x => x.Remove());

            // Remove review comments -> they are editorial annotations that are not useful in a
            // generated document, and a comment inside a loop would be cloned into several markers
            // sharing the same comment id, which corrupts the document (duplicate w:id).
            foreach (var comment in content.Descendants<CommentRangeStart>().ToList())
            {
                comment.RemoveWithEmptyParent();
            }
            foreach (var comment in content.Descendants<CommentRangeEnd>().ToList())
            {
                comment.RemoveWithEmptyParent();
            }
            foreach (var comment in content.Descendants<CommentReference>().ToList())
            {
                comment.RemoveWithEmptyParent();
            }

            // Bookmarks are intentionally kept so that cross-references (REF fields) and
            // internal hyperlinks (w:anchor) keep working in the generated document (issue #128).
            // They are sanitized after processing in MakeBookmarksValid, because loop expansion
            // can duplicate them and content removal can separate start/end pairs.

            // call extensions
            foreach (var extension in Context.Extensions)
            {
                extension.PreProcess(content);
            }
        }

        private void RemoveLineBreaksAroundSyntaxPatterns(IReadOnlyCollection<(PatternMatch, Text)> matches)
        {
            if (!Context.ProcessSettings.IgnoreLineBreaksAroundTags)
            {
                return;
            }
            static bool RemoveBreakAndCheckForText(OpenXmlElement openXmlElement)
            {
                if (openXmlElement is Break)
                {
                    openXmlElement.Remove();
                }
                return openXmlElement is Text;
            }
            foreach (var (_, text) in matches)
            {
                foreach (var next in text.ElementsSameLevelAfterInDocument())
                {
                    if (RemoveBreakAndCheckForText(next))
                    {
                        break;
                    }
                }
                foreach (var next in text.ElementsSameLevelBeforeInDocument())
                {
                    if (RemoveBreakAndCheckForText(next))
                    {
                        break;
                    }
                }
            }
        }

        private static IReadOnlyCollection<(PatternMatch, Text)> IsolateAndMergeTextTemplateMarkers(OpenXmlCompositeElement content)
        {
            var charMap = new CharacterMap(content);
            List<(PatternMatch, Text)> patternMatches = [];
            foreach (var m in PatternMatcher.FindSyntaxPatterns(charMap.Text))
            {
                var firstChar = charMap[m.Index];
                var lastChar = charMap[m.Index + m.Length - 1];
                // merge text creates or deletes elements but the index and the element with the match does not change
                // for this reason it does not matter that the new nodes are not in the charMap
                var mergedText = charMap.MergeText(firstChar, lastChar);
                mergedText.Element.Mark(m.Type);
                patternMatches.Add(new(m, mergedText.Element));
            }
            return patternMatches;
        }

        private static void Cleanup(OpenXmlCompositeElement element, bool removeEmptyElements)
        {
            InsertionPoint.RemoveAll(element);
            foreach (var markedText in element.Descendants<Text>().Where(x => x.IsMarked()).ToList())
            {
                var value = markedText.GetMarker();
                if (removeEmptyElements && value is not PatternType.Variable && value is not PatternType.Expression)
                {
                    var parent = markedText.Parent;
                    markedText.RemoveWithEmptyParent(preserveParagraphInCell: true);
                }
                else
                {
                    markedText.RemoveAttribute("mrk", null);
                }
            }

            // make dock properties ids unique
            uint id = 1;
            var dockProperties = element.Descendants<DocProperties>().ToList();
            var existingIds = new HashSet<uint>(dockProperties.Select(x => x.Id.Value).ToList());
            foreach (var docPropertiesWithSameId in dockProperties.GroupBy(x => x.Id).Where(x => x.Count() > 1))
            {
                foreach (var docProperties in docPropertiesWithSameId.Skip(1))
                {
                    while (existingIds.Contains(id))
                    {
                        id++;
                    }

                    docProperties.Id = id;
                    existingIds.Add(id);
                }
            }

            MakeBookmarksValid(element);

            //ensure all table cells have a paragraph
            // 'If a table cell does not include at least one block-level element, then this document shall be considered corrupt
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablecell?view=openxml-3.0.1#remarks
            foreach (var tableCell in element.Descendants<TableCell>())
            {
                if (!tableCell.ChildElements.OfType<Paragraph>().Any())
                {
                    tableCell.Append(new Paragraph());
                }
            }
        }

        /// <summary>
        /// Bookmarks (and their referencing REF fields / internal hyperlinks) are preserved during
        /// processing. However loop expansion clones bookmarks - producing duplicate ids and names -
        /// and removed content can separate a bookmarkStart from its bookmarkEnd. Word requires unique
        /// bookmark ids and names and matching start/end pairs, so this restores those invariants:
        /// orphaned starts/ends are removed, ids are renumbered uniquely and duplicate names are made unique.
        /// Bookmarks are referenced by name (not by id), so renumbering ids never breaks a reference.
        /// </summary>
        private static void MakeBookmarksValid(OpenXmlCompositeElement element)
        {
            var bookmarkStarts = element.Descendants<BookmarkStart>().ToList();
            var bookmarkEnds = element.Descendants<BookmarkEnd>().ToList();
            if (bookmarkStarts.Count == 0 && bookmarkEnds.Count == 0)
            {
                return;
            }

            // group ends by their (original) id in document order so the k-th start of an id
            // can be paired with the k-th end of the same id (loop clones keep start/end adjacent)
            var endsByOriginalId = new Dictionary<string, Queue<BookmarkEnd>>();
            foreach (var end in bookmarkEnds)
            {
                var key = end.Id?.Value ?? string.Empty;
                if (!endsByOriginalId.TryGetValue(key, out var queue))
                {
                    queue = new Queue<BookmarkEnd>();
                    endsByOriginalId[key] = queue;
                }
                queue.Enqueue(end);
            }

            var matchedEnds = new HashSet<BookmarkEnd>();
            var usedNames = new HashSet<string>(StringComparer.Ordinal);
            var nextId = 0;
            foreach (var start in bookmarkStarts)
            {
                var key = start.Id?.Value ?? string.Empty;
                if (!endsByOriginalId.TryGetValue(key, out var queue) || queue.Count == 0)
                {
                    // orphaned start without a matching end
                    start.RemoveWithEmptyParent();
                    continue;
                }

                var end = queue.Dequeue();
                matchedEnds.Add(end);

                var newId = nextId++.ToString(CultureInfo.InvariantCulture);
                start.Id = newId;
                end.Id = newId;

                // ensure unique names (a bookmark inside a loop is cloned with the same name)
                var name = start.Name?.Value ?? string.Empty;
                if (!usedNames.Add(name))
                {
                    string uniqueName;
                    var suffix = 1;
                    do
                    {
                        uniqueName = string.Concat(name, "_", suffix++.ToString(CultureInfo.InvariantCulture));
                    }
                    while (!usedNames.Add(uniqueName));
                    start.Name = uniqueName;
                }
            }

            // remove ends that were never matched to a start
            foreach (var end in bookmarkEnds)
            {
                if (!matchedEnds.Contains(end))
                {
                    end.RemoveWithEmptyParent();
                }
            }
        }

        private IReadOnlyCollection<ContentBlock> ExpandLoops(OpenXmlCompositeElement element, IReadOnlyCollection<(PatternMatch, Text)> matches)
        {
            Stack<ContentBlock> blockStack = new();
            blockStack.Push(new ContentBlock()); // dummy block for root
            foreach (var item in matches)
            {
                var match = item.Item1;
                var text = (Text)item.Item2;
                var patternType = match.Type;
                if (patternType is PatternType.InlineKeyWord)
                {
                    StartBlock(blockStack, match, patternType, text);
                    CloseBlock(blockStack, match, text);
                }

                if (patternType is PatternType.Condition or PatternType.CollectionStart or PatternType.IgnoreBlock or PatternType.Switch or PatternType.Case or PatternType.Default or PatternType.RangeStart)
                {
                    if (patternType is PatternType.Case or PatternType.Default)
                    {
                        AutoCloseCaseOrDefaultIfNeeded(blockStack, match, text);
                    }
                    StartBlock(blockStack, match, patternType, text);
                    StartBlock(blockStack, match, PatternType.None, text); // open the child content block of the loop or condition
                }
                else if (patternType is PatternType.ConditionElse or PatternType.CollectionSeparator)
                {
                    CloseBlock(blockStack, match, text);
                    StartBlock(blockStack, match, patternType, text);
                }
                if (patternType is PatternType.ConditionEnd or PatternType.CollectionEnd or PatternType.IgnoreEnd)
                {
                    CloseBlock(blockStack, match, text);
                    CloseBlock(blockStack, match, text);
                }
            }
            if (blockStack.Count != 1)
            {
                var notClosedBlocks = blockStack.Reverse().Skip(1).Select(x => x.StartMatch.Match.Value).Skip(1).ToList();
                throw new OpenXmlTemplateException($"Not all blocks are closed: {string.Join(", ", notClosedBlocks)}");
            }

            var rootBlock = blockStack.Peek();
            var rootChilds = rootBlock.ChildBlocks;

            foreach (var block in rootChilds)
            {
                block.AddInsertionPointsRecursively();
            }
#if DEBUG
            Console.WriteLine("--------- Assigned Insertion Points --------");
            Console.WriteLine(element.ToPrettyPrintXml());
#endif

            foreach (var block in rootChilds)
            {
                block.ExtractContentRecursively();
            }

#if DEBUG
            foreach (var block in rootChilds)
            {
                block.Validate();
            }
#endif

            return rootChilds;
        }

        private void StartBlock(Stack<ContentBlock> blockStack, PatternMatch match, PatternType value, Text text)
        {
            var newBlock = ContentBlock.Crate(Context, value, text, match);
            blockStack.Peek().AddChildBlock(newBlock);
            blockStack.Push(newBlock);
        }

        private static void AutoCloseCaseOrDefaultIfNeeded(Stack<ContentBlock> blockStack, PatternMatch match, Text text)
        {
            var current = blockStack.Peek();
            if (current.PatternType == PatternType.None && current.ParentBlock != null &&
                (current.ParentBlock.PatternType == PatternType.Case || current.ParentBlock.PatternType == PatternType.Default))
            {
                CloseBlock(blockStack, match, text); // Close None
                CloseBlock(blockStack, match, text); // Close Case/Default
            }
        }

        private static void CloseBlock(Stack<ContentBlock> blockStack, PatternMatch match, Text text)
        {
            if (blockStack.Count == 1)
            {
                throw new OpenXmlTemplateException($"Block was not open {text.InnerText}");
            }
            var closedBlock = blockStack.Pop();
            closedBlock.CloseBlock(text, match);
        }

        public void BindModel(string prefix, object model)
        {
            Context.ModelLookup.Add(prefix, model);
        }

        public void RegisterFormatter(IFormatter formatter)
        {
            if (formatter is IImageServiceProvider imageService)
            {
                Context.SetImageService(imageService.CreateImageService());
            }

            Context.VariableReplacer.RegisterFormatter(formatter);
        }

        public void RegisterExtension(ITemplateProcessorExtension extension)
        {
            Context.RegisterExtension(extension);
        }

        /// <summary>
        /// Runs the model-independent half of the pipeline: pre-process + isolate-text +
        /// block-tree-build, without expanding/rendering any block. Returns the top-level blocks.
        /// The document is mutated (marker attributes added, insertion points inserted, content
        /// extracted into blocks), but the result is exactly the state <see cref="RenderNode"/>
        /// expects. <see cref="DocxTemplate.GetTemplateSchema"/> caches the returned blocks so that a
        /// subsequent <see cref="DocxTemplate.Process"/> can render from them instead of rebuilding.
        /// </summary>
        internal IReadOnlyCollection<ContentBlock> BuildBlockTree(OpenXmlCompositeElement rootElement)
        {
#if DEBUG
            Console.WriteLine("----------- Original --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
            PreProcess(rootElement);
            var matches = IsolateAndMergeTextTemplateMarkers(rootElement);
            RemoveLineBreaksAroundSyntaxPatterns(matches);
#if DEBUG
            Console.WriteLine("----------- Isolate Texts --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
            return ExpandLoops(rootElement, matches);
        }
    }
}