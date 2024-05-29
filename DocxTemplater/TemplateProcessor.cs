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
        private readonly ModelLookup m_models;
        private readonly VariableReplacer m_variableReplacer;
        private readonly ScriptCompiler m_scriptCompiler;

        public ProcessSettings Settings { get; }

        private protected TemplateProcessor(
            ProcessSettings settings,
            ModelLookup modelLookup,
            VariableReplacer variableReplacer,
            ScriptCompiler scriptCompiler)
        {
            Settings = settings;
            m_models = modelLookup;
            m_variableReplacer = variableReplacer;
            m_scriptCompiler = scriptCompiler;
            m_variableReplacer.RegisterFormatter(new SubTemplateFormatter(modelLookup, settings));
        }

        public IReadOnlyDictionary<string, object> Models => m_models.Models;


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

        private IReadOnlyCollection<ContentBlock> ExpandLoops(OpenXmlCompositeElement element)
        {
            var syntaxTree = SyntaxTree.Build(element, m_variableReplacer, m_scriptCompiler);
            
            return syntaxTree.Roots;
        }

        internal static IReadOnlyCollection<OpenXmlElement> ExtractBlockContent(OpenXmlElement startText,
            OpenXmlElement endText, out OpenXmlElement leadingPart)
        {
            var commonParent = startText.FindCommonParent(endText) ??
                               throw new OpenXmlTemplateException("Start and end text are not in the same tree");
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

        public void BindModel(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        public void RegisterFormatter(IFormatter formatter)
        {
            m_variableReplacer.RegisterFormatter(formatter);
        }
    }
}