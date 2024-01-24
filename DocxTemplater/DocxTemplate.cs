using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Blocks;
using DocxTemplater.Formatter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ContentBlock = DocxTemplater.Blocks.ContentBlock;

namespace DocxTemplater
{
    public sealed class DocxTemplate : IDisposable
    {
        private readonly Stream m_stream;
        private readonly WordprocessingDocument m_wpDocument;
        private readonly ModelDictionary m_models;

        private static readonly FileFormatVersions TargetMinimumVersion = FileFormatVersions.Office2010;


        private readonly VariableReplacer m_variableReplacer;
        private readonly ScriptCompiler m_scriptCompiler;

        public DocxTemplate(Stream docXStream, ProcessSettings settings = null)
        {
            ArgumentNullException.ThrowIfNull(docXStream);
            Settings = settings ?? ProcessSettings.Default;
            m_stream = new MemoryStream();
            docXStream.CopyTo(m_stream);
            m_stream.Position = 0;

            // Add the MarkupCompatibilityProcessSettings
            OpenSettings openSettings = new()
            {
                MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    TargetMinimumVersion)
            };

            m_wpDocument = WordprocessingDocument.Open(m_stream, true, openSettings);
            m_models = new ModelDictionary();
            m_scriptCompiler = new ScriptCompiler(m_models);
            m_variableReplacer = new VariableReplacer(m_models, Settings);
            Processed = false;
        }

        public ProcessSettings Settings { get; }

        public static DocxTemplate Open(string pathToTemplate, ProcessSettings settings = null)
        {
            using var fileStream = new FileStream(pathToTemplate, FileMode.Open, FileAccess.Read);
            return new DocxTemplate(fileStream, settings);
        }

        public IReadOnlyDictionary<string, object> Models => m_models.Models;

        public bool Processed { get; private set; }

        public void RegisterFormatter(IFormatter formatter)
        {
            m_variableReplacer.RegisterFormatter(formatter);
        }

        public void BindModel(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        public void Save(string targetPath)
        {
            using var fileStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write);
            Save(fileStream);
        }

        public void Save(Stream target)
        {
            Process().CopyTo(target);
        }

        public Stream AsStream()
        {
            return Process();
        }

        public void Validate()
        {
            var v = new OpenXmlValidator(TargetMinimumVersion);
            var errs = v.Validate(m_wpDocument);
            if (errs.Any())
            {
                var sb = new System.Text.StringBuilder();
                foreach (var err in errs)
                {
                    sb.AppendLine($"{err.Description} Node: {err.Node} RelatedNode: {err.RelatedNode}");
                }
                throw new OpenXmlTemplateException($"Validation failed - {sb}");
            }
        }

        public Stream Process()
        {
            if (m_wpDocument.MainDocumentPart == null || Processed)
            {
                m_stream.Position = 0;
                return m_stream;
            }
            Processed = true;
            foreach (var header in m_wpDocument.MainDocumentPart.HeaderParts)
            {
                ProcessNode(header.Header);
            }
            ProcessNode(m_wpDocument.MainDocumentPart.RootElement);
            foreach (var footer in m_wpDocument.MainDocumentPart.FooterParts)
            {
                ProcessNode(footer.Footer);
            }
            m_wpDocument.Save();
            m_stream.Position = 0;
            return m_stream;
        }

        private void ProcessNode(OpenXmlPartRootElement rootElement)
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
            m_variableReplacer.ReplaceVariables(rootElement);
            foreach (var loop in loops)
            {
                loop.Expand(m_models, rootElement);
            }
            Cleanup(rootElement);
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
                var firstText = (Text)firstChar.Element;
                var lastText = (Text)lastChar.Element;
                var mergedText = firstText.MergeText(firstChar.Index, lastText, m.Length);
                mergedText.Mark(m.Type);
                // TODO: Ist this possible without recreate charMap?
                charMap.MarkAsDirty();
            }
        }

        private static void Cleanup(OpenXmlCompositeElement element)
        {
            InsertionPoint.RemoveAll(element);
            foreach (var markedText in element.Descendants<Text>().Where(x => x.IsMarked()).ToList())
            {
                var value = markedText.GetMarker();
                if (value is PatternType.CollectionStart or PatternType.CollectionEnd or PatternType.ConditionEnd or PatternType.ConditionElse)
                {
                    var parent = markedText.Parent;
                    markedText.RemoveWithEmptyParent();
                    if (parent != null && parent.ChildElements.All(x => x is Languages))
                    {
                        parent.RemoveWithEmptyParent();
                    }
                }
                else
                {
                    markedText.RemoveAttribute("mrk", null);
                }
            }

            // make all Bookmark ids unique
            int id = 0;
            foreach (var bookmarkStart in element.Descendants<BookmarkStart>())
            {
                bookmarkStart.Id = $"{id++}";
                bookmarkStart.NextSibling<BookmarkEnd>().Id = bookmarkStart.Id;
            }
        }

        private IReadOnlyCollection<ContentBlock> ExpandLoops(OpenXmlPartRootElement element)
        {

            // TODO: store metadata for tag in cache
            var blockStack = new Stack<(ContentBlock Block, PatternMatch Match, Text MatchedTextNode)>();
            blockStack.Push((new ContentBlock(m_variableReplacer), null, null)); // dummy block for root
            // find all begin or end markers
            foreach (var text in element.Descendants<Text>().Where(x => x.IsMarked()))
            {
                var value = text.GetMarker();
                if (value is PatternType.CollectionStart)
                {
                    var match = PatternMatcher.FindSyntaxPatterns(text.Text).Single();
                    if (match.Formatter.Equals("dyntable", StringComparison.InvariantCultureIgnoreCase))
                    {
                        blockStack.Push((new DynamicTableBlock(match.Variable, m_variableReplacer), match, text));
                    }
                    else
                    {
                        blockStack.Push((new LoopBlock(match.Variable, m_variableReplacer), match, text));
                    }
                }
                else if (value == PatternType.Condition)
                {
                    var match = PatternMatcher.FindSyntaxPatterns(text.Text).Single();
                    blockStack.Push((new ConditionalBlock(match.Condition, m_variableReplacer, m_scriptCompiler), match, text));
                }
                else if (value == PatternType.ConditionElse)
                {
                    var (block, patternMatch, matchedTextNode) = blockStack.Pop();
                    if (block is not ConditionalBlock)
                    {
                        throw new OpenXmlTemplateException($"'{block}' is not closed");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    block.SetContent(leadingPart, loopContent);
                    blockStack.Push((block, patternMatch, text)); // push same block again on Stack but with other text element
                }
                else if (value == PatternType.ConditionEnd)
                {
                    var (block, _, matchedTextNode) = blockStack.Pop();
                    if (block is not ConditionalBlock)
                    {
                        throw new OpenXmlTemplateException($"'{block}' is not closed");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    block.SetContent(leadingPart, loopContent);
                    blockStack.Peek().Block.AddInnerBlock(block);
                }
                else if (value == PatternType.CollectionEnd)
                {
                    var (block, patternMatch, matchedTextNode) = blockStack.Pop();
                    if (patternMatch.Type != PatternType.CollectionStart)
                    {
                        throw new OpenXmlTemplateException($"'{block}' is not closed");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    block.SetContent(leadingPart, loopContent);
                    blockStack.Peek().Block.AddInnerBlock(block);
                }

            }
            var (contentBlock, _, _) = blockStack.Pop();
            return contentBlock.ChildBlocks;
        }


        internal static IReadOnlyCollection<OpenXmlElement> ExtractBlockContent(OpenXmlElement startText, OpenXmlElement endText, out OpenXmlElement leadingPart)
        {
            var commonParent = startText.FindCommonParent(endText) ?? throw new OpenXmlTemplateException("Start and end text are not in the same tree");
            var result = new List<OpenXmlElement>();
            if (commonParent is TableRow)
            {
                var previousRow = commonParent.PreviousSibling();
                if (previousRow == null)
                {
                    commonParent.InsertBeforeSelf(new TableRow());
                }
                leadingPart = commonParent.PreviousSibling();
                commonParent.Remove();
                result.Add(commonParent);
            }
            else
            {
                // find childs of common parent that contains start and end text
                var startChildOfCommonParent = commonParent.ChildElements.Single(c =>
                    c == startText || c.Descendants<Text>().Any(d => d == startText));
                var endChildOfCommonParent =
                    commonParent.ChildElements.Single(c =>
                        c == endText || c.Descendants<Text>().Any(d => d == endText));

                var startSplit = startChildOfCommonParent.SplitAfterElement(startText);
                leadingPart = startSplit.First();
                if (startChildOfCommonParent == endChildOfCommonParent)
                {
                    result.AddRange(commonParent.ChildsBetween(startSplit.First(), endChildOfCommonParent).ToList());
                }
                else
                {
                    var endSplit = endChildOfCommonParent.SplitBeforeElement(endText);
                    result.AddRange(commonParent.ChildsBetween(leadingPart, endSplit.Last()).ToList());
                }

                foreach (var element in result)
                {
                    element.Remove();
                }
            }
            return result;
        }


        public void Dispose()
        {
            m_stream?.Dispose();
            m_wpDocument?.Dispose();
        }
    }
}
