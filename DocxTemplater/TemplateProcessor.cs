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

            DocxTemplate.IsolateAndMergeTextTemplateMarkers(rootElement);

#if DEBUG
            Console.WriteLine("----------- Isolate Texts --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif

            var loops = ExpandLoops(rootElement);
#if DEBUG
            Console.WriteLine("----------- After Loops --------");
            Console.WriteLine(rootElement.ToPrettyPrintXml());
#endif
            Context.VariableReplacer.ReplaceVariables(rootElement, Context);
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

        private static void PreProcess(OpenXmlCompositeElement content)
        {
            content.Descendants<ProofError>().ToList().ForEach(x => x.Remove());
        }

        private static void IsolateAndMergeTextTemplateMarkers(OpenXmlCompositeElement content)
        {
            var charMap = new CharacterMap(content);
            foreach (var m in PatternMatcher.FindSyntaxPatterns(charMap.Text))
            {
                var firstChar = charMap[m.Index];
                var lastChar = charMap[m.Index + m.Length - 1];
                var firstText = firstChar.Element;
                var lastText = lastChar.Element;
                var mergedText = firstText.MergeText(firstChar.CharIndexInText, lastText, m.Length);
                mergedText.Mark(m.Type);
                // TODO: Ist this possible without recreate charMap?
                charMap.Recreate();
            }
        }

        private static void Cleanup(OpenXmlCompositeElement element, bool removeEmptyElements)
        {
            InsertionPoint.RemoveAll(element);
            foreach (var markedText in element.Descendants<Text>().Where(x => x.IsMarked()).ToList())
            {
                var value = markedText.GetMarker();
                if (removeEmptyElements && value is not PatternType.Variable)
                {
                    var parent = markedText.Parent;
                    markedText.RemoveWithEmptyParent();
                }
                else
                {
                    markedText.RemoveAttribute("mrk", null);
                }
            }

            // remove all bookmarks -> not useful for generated documents and complex to handle
            // because of special cases in tables see
            // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bookmarkstart?view=openxml-3.0.1#remarks
            foreach (var bookmark in element.Descendants<BookmarkStart>().ToList())
            {
                bookmark.RemoveWithEmptyParent();
            }
            foreach (var bookmark in element.Descendants<BookmarkEnd>().ToList())
            {
                bookmark.RemoveWithEmptyParent();
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

        private IReadOnlyCollection<ContentBlock> ExpandLoops(OpenXmlCompositeElement element)
        {
            Stack<ContentBlock> blockStack = new();
            blockStack.Push(new ContentBlock()); // dummy block for root
            foreach (var text in element.Descendants<Text>().ToList().Where(x => x.IsMarked()))
            {
                var value = text.GetMarker();
                var match = PatternMatcher.FindSyntaxPatterns(text.Text).Single();

                if (value is PatternType.InlineKeyWord)
                {
                    StartBlock(blockStack, match, value, text);
                    CloseBlock(blockStack, match, text);
                }

                if (value is PatternType.Condition or PatternType.CollectionStart)
                {
                    StartBlock(blockStack, match, value, text);
                    StartBlock(blockStack, match, PatternType.None, text);
                }
                else if (value is PatternType.ConditionElse or PatternType.CollectionSeparator)
                {
                    CloseBlock(blockStack, match, text);
                    StartBlock(blockStack, match, value, text);
                }
                if (value is PatternType.ConditionEnd or PatternType.CollectionEnd)
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
    }
}