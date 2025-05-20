#if DEBUG
using System;
#endif
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Blocks;
using DocxTemplater.Formatter;
using System.Collections.Generic;
using System.Linq;
using DocxTemplater.Extensions;

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
#if DEBUG
            Console.WriteLine("----------- Original --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
            PreProcess(rootElement);

            var matches = DocxTemplate.IsolateAndMergeTextTemplateMarkers(rootElement);

            RemoveLineBreaksAroundSyntaxPatterns(matches);

#if DEBUG
            Console.WriteLine("----------- Isolate Texts --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif

            var loops = ExpandLoops(rootElement, matches);
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

            Cleanup(rootElement, Context, removeEmptyElements: true);
#if DEBUG
            Console.WriteLine("----------- Completed --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
        }

        private void PreProcess(OpenXmlCompositeElement content)
        {
            // remove spell check 'ProofError' elements
            content.Descendants<ProofError>().ToList().ForEach(x => x.Remove());

            // remove all bookmarks -> not useful for generated documents and complex to handle
            // because of special cases in tables see
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bookmarkstart?view=openxml-3.0.1#remarks
            foreach (var bookmark in content.Descendants<BookmarkStart>().ToList())
            {
                bookmark.RemoveWithEmptyParent();
            }
            foreach (var bookmark in content.Descendants<BookmarkEnd>().ToList())
            {
                bookmark.RemoveWithEmptyParent();
            }

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

        private static void Cleanup(OpenXmlCompositeElement element, ITemplateProcessingContextAccess context, bool removeEmptyElements)
        {
            InsertionPoint.RemoveAll(element);
            
            // Get all paragraphs that contain marked text
            var paragraphsWithMarkedText = element.Descendants<Text>()
                .Where(x => x.IsMarked())
                .Select(x => x.GetFirstAncestor<Paragraph>())
                .Where(p => p != null)
                .Distinct()
                .ToList();
                
            // Mark paragraphs that only contain blocks - if a paragraph only contains marked text
            foreach (var paragraph in paragraphsWithMarkedText)
            {
                bool containsOnlyMarkedText = true;
                foreach (var textElement in paragraph.Descendants<Text>())
                {
                    // If there's any unmarked text or variable text, the paragraph doesn't contain only blocks
                    if (!textElement.IsMarked() || textElement.GetMarker() == PatternType.Variable)
                    {
                        containsOnlyMarkedText = false;
                        break;
                    }
                }
                
                if (containsOnlyMarkedText)
                {
                    paragraph.MarkAsContainingOnlyBlocks();
                }
            }
            
            // Process marked text
            foreach (var markedText in element.Descendants<Text>().Where(x => x.IsMarked()).ToList())
            {
                var value = markedText.GetMarker();
                
                // If it's a closing tag (end tag), mark its paragraph for potential removal
                if (value is PatternType.ConditionEnd or PatternType.CollectionEnd or PatternType.IgnoreEnd)
                {
                    var paragraph = markedText.GetFirstAncestor<Paragraph>();
                    if (paragraph != null)
                    {
                        paragraph.MarkAsContainingOnlyBlocks();
                    }
                }
                
                if (removeEmptyElements && value is not PatternType.Variable)
                {
                    var parent = markedText.Parent;
                    markedText.RemoveWithEmptyParent();
                }
                else
                {
                    markedText.RemoveMarker();
                }
            }
            
            // Handle paragraphs marked as containing only blocks
            if (context?.ProcessSettings?.RemoveParagraphsContainingOnlyBlocks == true)
            {
                // First pass: remove paragraphs explicitly marked as containing only blocks
                foreach (var paragraph in element.Descendants<Paragraph>()
                    .Where(p => p.IsMarkedAsContainingOnlyBlocks())
                    .ToList())
                {
                    // Check if the paragraph has any actual content after processing
                    bool isEmpty = true;
                    
                    // Skip checking paragraph properties
                    foreach (var child in paragraph.ChildElements)
                    {
                        if (child is ParagraphProperties)
                            continue;
                            
                        // If it has any content elements that are not properties
                        if (child is Run run)
                        {
                            // Check if run has any non-empty text
                            if (run.ChildElements.Any(c => 
                                (c is Text text && !string.IsNullOrWhiteSpace(text.Text))))
                            {
                                isEmpty = false;
                                break;
                            }
                        }
                        else
                        {
                            // If there's any non-property element
                            isEmpty = false;
                            break;
                        }
                    }
                    
                    if (isEmpty)
                    {
                        paragraph.Remove();
                    }
                }
                
                // Second pass: Check for empty paragraphs with just IpId attributes
                // These are usually left after insertion points for blocks are processed
                foreach (var paragraph in element.Descendants<Paragraph>().ToList())
                {
                    // Fix GetAttribute to use try/catch to avoid exception
                    bool hasIpId = false;
                    try 
                    {
                        hasIpId = paragraph.GetAttribute("IpId", null) != null;
                    }
                    catch (KeyNotFoundException)
                    {
                        // Attribute not allowed on this element, skip it
                        continue;
                    }
                    
                    // If there's an IpId attribute and no visible content
                    if (hasIpId && 
                        !paragraph.Descendants<Text>().Any(t => !string.IsNullOrWhiteSpace(t.Text)))
                    {
                        paragraph.Remove();
                    }
                }
                
                // Third pass: Look for any remaining completely empty paragraphs
                foreach (var paragraph in element.Descendants<Paragraph>().ToList())
                {
                    bool isEmpty = true;
                    
                    // Skip checking paragraph properties
                    foreach (var child in paragraph.ChildElements)
                    {
                        if (child is ParagraphProperties)
                            continue;
                            
                        // If it has any content elements that are not properties
                        if (child is Run run)
                        {
                            // Check if run has any non-empty text or other elements that aren't properties
                            if (run.ChildElements.Any(c => 
                                (c is Text text && !string.IsNullOrWhiteSpace(text.Text)) || 
                                !(c is RunProperties)))
                            {
                                isEmpty = false;
                                break;
                            }
                        }
                        else
                        {
                            // If there's any non-property element
                            isEmpty = false;
                            break;
                        }
                    }
                    
                    if (isEmpty)
                    {
                        paragraph.Remove();
                    }
                }
            }
            
            // Clean up any remaining marks
            foreach (var paragraph in element.Descendants<Paragraph>()
                .Where(p => p.IsMarkedAsContainingOnlyBlocks())
                .ToList())
            {
                paragraph.RemoveContainsOnlyBlocksMarker();
            }
            
            // Clean up all markers from the dictionary
            ParagraphExtensions.CleanupAllMarkers();
            
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

                if (patternType is PatternType.Condition or PatternType.CollectionStart or PatternType.IgnoreBlock)
                {
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
            Context.VariableReplacer.RegisterFormatter(formatter);
        }

        public void RegisterExtension(ITemplateProcessorExtension extension)
        {
            Context.RegisterExtension(extension);
        }
    }
}