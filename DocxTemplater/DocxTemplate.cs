﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Blocks;
using DocxTemplater.Formatter;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace DocxTemplater
{
    public sealed class DocxTemplate : IDisposable
    {
        private readonly Stream m_stream;
        private readonly WordprocessingDocument m_wpDocument;
        private readonly ModelLookup m_models;

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
            m_models = new ModelLookup();
            m_scriptCompiler = new ScriptCompiler(m_models, Settings);
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
                var firstText = (Text)firstChar.Element;
                var lastText = (Text)lastChar.Element;
                var mergedText = firstText.MergeText(firstChar.Index, lastText, m.Length);
                mergedText.Mark(m.Type);
                // TODO: Ist this possible without recreate charMap?
                charMap.MarkAsDirty();
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

            // make all Bookmark ids unique
            uint id = 0;
            foreach (var bookmarkStart in element.Descendants<BookmarkStart>())
            {
                bookmarkStart.Id = $"{id++}";
                bookmarkStart.NextSibling<BookmarkEnd>().Id = bookmarkStart.Id;
            }

            // make dock properties ids unique
            id = 1;
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

        private IReadOnlyCollection<ContentBlock> ExpandLoops(OpenXmlPartRootElement element)
        {

            // TODO: store metadata for tag in cache
            var blockStack = new Stack<(ContentBlock Block, PatternType type, Text MatchedTextNode)>();
            blockStack.Push((new ContentBlock(m_variableReplacer), PatternType.None, null)); // dummy block for root
            // find all begin or end markers
            foreach (var text in element.Descendants<Text>().ToList().Where(x => x.IsMarked()))
            {
                var value = text.GetMarker();
                if (value is PatternType.CollectionStart)
                {
                    var match = PatternMatcher.FindSyntaxPatterns(text.Text).Single();
                    if (match.Formatter.Equals("dyntable", StringComparison.InvariantCultureIgnoreCase))
                    {
                        blockStack.Push((new DynamicTableBlock(match.Variable, m_variableReplacer), value, text));
                    }
                    else
                    {
                        blockStack.Push((new LoopBlock(match.Variable, m_variableReplacer), value, text));
                    }
                }
                else if (value == PatternType.CollectionSeparator)
                {
                    var (block, _, matchedTextNode) = blockStack.Pop();
                    if (block is not LoopBlock collectionStartBlock)
                    {
                        throw new OpenXmlTemplateException($"Separator in '{block}' is invalid");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    var insertPoint = InsertionPoint.CreateForElement(leadingPart);
                    collectionStartBlock.SetContent(insertPoint, loopContent);
                    var separatorBlock = new ContentBlock(m_variableReplacer, collectionStartBlock, insertPoint);
                    collectionStartBlock.SetSeparatorBlock(separatorBlock);
                    blockStack.Push((separatorBlock, value, text));
                }
                else if (value == PatternType.CollectionEnd)
                {
                    var (block, startType, matchedTextNode) = blockStack.Pop();
                    if (startType is not PatternType.CollectionStart and not PatternType.CollectionSeparator)
                    {
                        throw new OpenXmlTemplateException($"'{text.InnerText}' is missing collection start: {text.ElementBeforeInDocument<Text>()?.InnerText} >> {text.InnerText} << {text.ElementAfterInDocument<Text>()?.InnerText}");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    block.SetContent(InsertionPoint.CreateForElement(leadingPart), loopContent);
                    blockStack.Peek().Block.AddInnerBlock(block.RootBlock);
                }
                else if (value == PatternType.Condition)
                {
                    var match = PatternMatcher.FindSyntaxPatterns(text.Text).Single();
                    blockStack.Push((new ConditionalBlock(match.Condition, m_variableReplacer, m_scriptCompiler), value, text));
                }
                else if (value == PatternType.ConditionElse)
                {
                    var (block, startType, matchedTextNode) = blockStack.Pop();
                    if (block is not ConditionalBlock conditionalBlock)
                    {
                        throw new OpenXmlTemplateException($"else block in '{block}' is invalid");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    var insertPoint = InsertionPoint.CreateForElement(leadingPart);
                    conditionalBlock.SetContent(insertPoint, loopContent);
                    var elseBlock = new ContentBlock(m_variableReplacer, conditionalBlock, insertPoint);
                    conditionalBlock.SetElseBlock(elseBlock);
                    blockStack.Push((elseBlock, value, text)); // push else block on stack but with other text element

                }
                else if (value == PatternType.ConditionEnd)
                {
                    var (block, startType, matchedTextNode) = blockStack.Pop();
                    if (startType is not PatternType.Condition and not PatternType.ConditionElse)
                    {
                        throw new OpenXmlTemplateException($"'{text.InnerText}' is missing condition start: {text.ElementBeforeInDocument<Text>()?.InnerText} >> {text.InnerText} << {text.ElementAfterInDocument<Text>()?.InnerText}");
                    }
                    var loopContent = ExtractBlockContent(matchedTextNode, text, out var leadingPart);
                    var insertPoint = InsertionPoint.CreateForElement(leadingPart);
                    block.SetContent(insertPoint, loopContent);
                    blockStack.Peek().Block.AddInnerBlock(block.RootBlock);
                }
            }
            if (blockStack.Count != 1)
            {
                var notClosedBlocks = blockStack.Reverse().Select(x => x.Block).Skip(1).ToList();
                throw new OpenXmlTemplateException($"Not all blocks are closed: {string.Join(", ", notClosedBlocks)}");
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
