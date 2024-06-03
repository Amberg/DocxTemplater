using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Formatter;
using System;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;

namespace DocxTemplater.Blocks
{
    internal class ContentBlock
    {
        protected InsertionPoint m_insertionPoint;
        protected IReadOnlyCollection<OpenXmlElement> m_content;
        protected readonly List<ContentBlock> m_childBlocks;
        protected readonly VariableReplacer m_variableReplacer;
        private InsertionPoint m_lastElementMarker;

        public ContentBlock()
            :this(null, PatternType.None, null, null)
        {}

        public ContentBlock(VariableReplacer variableReplacer, PatternType patternType, Text startTextNode, PatternMatch startMatch)
        {
            m_content = new List<OpenXmlElement>();
            m_childBlocks = new List<ContentBlock>();
            m_variableReplacer = variableReplacer;
            PatternType = patternType;
            StartTextNode = startTextNode;
            StartMatch = startMatch;
        }

        public static ContentBlock Crate(
            VariableReplacer variableReplacer,
            ScriptCompiler scriptCompiler,
            PatternType patternType,
            Text startTextNode,
            PatternMatch matchedStartNode)
        {
            return patternType switch
            {
                PatternType.CollectionStart when matchedStartNode.Formatter.Equals("dyntable", StringComparison.InvariantCultureIgnoreCase) => new DynamicTableBlock(variableReplacer, patternType, startTextNode, matchedStartNode),
                PatternType.CollectionStart => new LoopBlock(variableReplacer, patternType, startTextNode, matchedStartNode),
                PatternType.CollectionSeparator => new CollectionSeparatorBlock(variableReplacer, patternType, startTextNode, matchedStartNode),
                PatternType.Condition => new ConditionalBlock(variableReplacer, scriptCompiler, patternType, startTextNode, matchedStartNode),
                _ => new ContentBlock(variableReplacer, patternType, startTextNode, matchedStartNode)
            };
        }

        public PatternType PatternType
        {
            get;
        }

        public Text StartTextNode
        {
            get;
        }

        public Text EndTextNode
        {
            get;
            private set;
        }

        public PatternMatch StartMatch
        {
            get;
        }

        public PatternMatch EndMatch
        {
            get;
            private set;
        }

        public IReadOnlyCollection<OpenXmlElement> Content => m_content;

        public ContentBlock ParentBlock { get; private set; }

        public OpenXmlElement ParentNode { get; set; }

        public OpenXmlElement LastElement { get; set; }

        public OpenXmlElement FirstElement { get; set; }


        public IReadOnlyCollection<ContentBlock> ChildBlocks => m_childBlocks;

        public virtual void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            InsertContentAndReplaceVariables(models, parentNode);
            ExpandChildBlocks(models, parentNode);
            RemoveChildBlockInsertionPoints(parentNode);
        }

        protected virtual void InsertContentAndReplaceVariables(ModelLookup models, OpenXmlElement parentNode)
        {
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned);
            Console.WriteLine($"-------------------- After Insert Content {this} -----------------");
            Console.WriteLine(parentNode.ToPrettyPrintXml());
            m_variableReplacer.ReplaceVariables(cloned);
        }

        public virtual void ExpandChildBlocks(ModelLookup models, OpenXmlElement parentNode)
        {
            foreach (var child in m_childBlocks)
            {
                child.Expand(models, parentNode);
            }
        }

        protected virtual void RemoveChildBlockInsertionPoints(OpenXmlElement parentNode)
        {
            foreach (var child in m_childBlocks)
            {
                child.RemoveAnchor(parentNode);
            }
        }

        protected virtual void InsertContent(OpenXmlElement parentNode, IEnumerable<OpenXmlElement> paragraphs)
        {
            if (m_insertionPoint.Id == "Condition_4")
            {
                Console.WriteLine("adding");
            }
            var element = m_insertionPoint.GetElement(parentNode) ?? throw new OpenXmlTemplateException($"Insertion point {m_insertionPoint.Id} not found");
            element.InsertAfterSelf(paragraphs);
        }

        public void RemoveAnchor(OpenXmlElement parentNode)
        {
            if (m_insertionPoint.Id == "Condition_4")
            {
                Console.WriteLine("Removing anchor");
            }
            var element = m_insertionPoint.GetElement(parentNode) ?? throw new OpenXmlTemplateException($"Insertion point {m_insertionPoint.Id} not found");
            element.Remove();
        }

        public override string ToString()
        {
            return $"{StartTextNode.Text}_{m_insertionPoint}";
        }


        public virtual void AddChildBlock(ContentBlock block)
        {
            block.ParentBlock = this;
            m_childBlocks.Add(block);
        }

        public void Print(int i)
        {
            Console.WriteLine(new string('-', i + 1) + ToString());
            Console.WriteLine(m_content.ToPrettyPrintXml());
            foreach (var block in m_childBlocks)
            {
                block.Print(i + 2);
            }
        }

        public void CloseBlock(Text endTextNode, PatternMatch matchedEndNode)
        {
            EndTextNode = endTextNode;
            EndMatch = matchedEndNode;
        }

        public virtual void ExtractContentRecursively()
        {
            foreach (var child in m_childBlocks)
            {
                child.ExtractContentRecursively();
            }
            m_content = ParentNode.ChildsBetween(FirstElement, LastElement).ToList();
            foreach (var content in m_content)
            {
                content.Remove();
            }
        }

        public virtual void Validate()
        {
            foreach (var child in ChildBlocks)
            {
                var childIp = child.m_insertionPoint;
                if (!m_content.SelectMany(x => x.Descendants()).Concat(m_content).Any(childIp.IsForElement))
                {
                    throw new OpenXmlTemplateException($"Insertion Point {childIp.Id} of child {child} not found in {this}\r\n{m_content.ToPrettyPrintXml()}");
                }
            }
        }

        public void AddInsertionPointsRecursively()
        {
            AddInsertionPoints();
            foreach (var child in m_childBlocks)
            {
                child.AddInsertionPointsRecursively();
            }
        }

        private void AddInsertionPoints()
        {
            ParentNode = StartTextNode.FindCommonParent(EndTextNode) ?? throw new OpenXmlTemplateException("Start and end text are not in the same tree");
            if (ParentNode is TableRow tableRow)
            {
                FirstElement = tableRow.InsertBeforeSelf(new TableRow());
                LastElement = tableRow.InsertAfterSelf(new TableRow());
                ParentNode = tableRow.Parent;
            }
            else if (ParentNode is Table)
            {
                var firstRow = ParentNode.Elements<TableRow>().First(StartTextNode.IsChildOf);
                var lastRow = ParentNode.Elements<TableRow>().First(EndTextNode.IsChildOf);
                FirstElement = firstRow.InsertBeforeSelf(new TableRow());
                LastElement = lastRow.InsertAfterSelf(new TableRow());
            }
            else
            {
                // find childs of common parent that contains start and end text
                var startChildOfCommonParent = ParentNode.ChildElements.Single(c => c == StartTextNode || c.Descendants<Text>().Any(d => d == StartTextNode));
                var endChildOfCommonParent = ParentNode.ChildElements.Single(c => c == EndTextNode || c.Descendants<Text>().Any(d => d == EndTextNode));
                var split = startChildOfCommonParent.SplitAfterElement(StartTextNode);
                OpenXmlElement anchorElement = null;
                if (split.Count == 1)
                {
                    // already split - only first part returned
                    var nextElement = split.First().NextSibling();

                    // if two blocks opens there is already an anchor of the parent element
                    while (InsertionPoint.HasAlreadyInsertionPointMarker(nextElement))
                    {
                        nextElement = nextElement.NextSibling();
                    }
                    anchorElement = nextElement.InsertBeforeSelf(new Paragraph());
                }
                else
                {
                    var splitLastPart = split.Last();
                    anchorElement = splitLastPart.InsertBeforeSelf(new Paragraph());
                }

                if (startChildOfCommonParent == endChildOfCommonParent)
                {
                    FirstElement = anchorElement;
                    LastElement = endChildOfCommonParent;
                }
                else
                {
                    var endSplit = endChildOfCommonParent.SplitBeforeElement(EndTextNode);
                    FirstElement = anchorElement;
                    LastElement = endSplit.Last();
                }
            }
            m_insertionPoint = InsertionPoint.CreateForElement(FirstElement, $"{PatternType}");
            m_lastElementMarker = InsertionPoint.CreateForElement(LastElement, $"End_{PatternType}");
        }
    }
}