using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;
using DocxTemplater.Blocks;
using System;
using DocxTemplater.Formatter;

namespace DocxTemplater
{
    internal class SyntaxTree
    {
        private readonly OpenXmlCompositeElement m_element;
        private readonly VariableReplacer m_variableReplacer;
        private readonly ScriptCompiler m_scriptCompiler;
        private readonly Stack<SyntaxBlock> m_blockStack;

        private SyntaxTree(OpenXmlCompositeElement element, VariableReplacer variableReplacer, ScriptCompiler scriptCompiler)
        {
            m_element = element;
            m_variableReplacer = variableReplacer;
            m_scriptCompiler = scriptCompiler;
            m_blockStack = new Stack<SyntaxBlock>();
        }

        public IReadOnlyCollection<ContentBlock> Roots => m_blockStack.Peek().ContentBlock.ChildBlocks.ToList();

        public static SyntaxTree Build(OpenXmlCompositeElement element, VariableReplacer variableReplacer,
            ScriptCompiler scriptCompiler)
        {
            var syntaxTree = new SyntaxTree(element, variableReplacer, scriptCompiler);
            syntaxTree.Build();
            return syntaxTree;
        }

        private void Build()
        {
            m_blockStack.Push(new SyntaxBlock() { PatternType = PatternType.None, ContentBlock = new ContentBlock(m_variableReplacer) }); // dummy block for root
            foreach (var text in m_element.Descendants<Text>().ToList().Where(x => x.IsMarked()))
            {
                var value = text.GetMarker();
                var match = PatternMatcher.FindSyntaxPatterns(text.Text).Single();

                if (value is PatternType.CollectionStart)
                {
                    StartBlock(match, value, text, m_blockStack);
                }
                else if (value is PatternType.Condition)
                {
                    StartBlock(match, value, text, m_blockStack);
                    StartBlock(match, PatternType.None, text, m_blockStack);
                }
                else if (value is PatternType.ConditionElse)
                {
                    CloseBlock(match, text);
                    StartBlock(match, value, text, m_blockStack);
                }
                else if (value is PatternType.CollectionSeparator)
                {
                    StartBlock(match, value, text, m_blockStack);
                }
                if (value is PatternType.ConditionEnd)
                {
                    CloseBlock(match, text);
                    CloseBlock(match, text);
                }
                if (value is PatternType.CollectionEnd)
                {
                    var currentBlock = m_blockStack.Peek().PatternType;
                    if (currentBlock is PatternType.CollectionSeparator)
                    {
                        CloseBlock(match, text);
                    }
                    CloseBlock(match, text);
                }
            }

            if (m_blockStack.Count != 1)
            {
                var notClosedBlocks = m_blockStack.Reverse().Select(x => x.MatchedStartNode.Match.Value).Skip(1).ToList();
                throw new OpenXmlTemplateException($"Not all blocks are closed: {string.Join(", ", notClosedBlocks)}");
            }

            var rootBlock = m_blockStack.Peek();
            foreach (var block in rootBlock.TraverseUp())
            {
                if (block == rootBlock)
                {
                    continue;
                }
                ExtractBlockContent(block);
            }
            rootBlock.AssignBlockChildren();
        }

        private void CloseBlock(PatternMatch match, Text text)
        {
            var closedBlock = m_blockStack.Pop();
            closedBlock.MatchedEndNode = match;
            closedBlock.EndTextNode = text;
            var commonParent = closedBlock.StartTextNode.FindCommonParent(closedBlock.EndTextNode) ?? throw new OpenXmlTemplateException("Start and end text are not in the same tree");
            closedBlock.ParentNode = commonParent;
        }

        private static void StartBlock(PatternMatch match, PatternType value, Text text, Stack<SyntaxBlock> blockStack)
        {
            var newBlock = new SyntaxBlock()
            {
                MatchedStartNode = match,
                PatternType = value,
                StartTextNode = text
            };
            blockStack.Peek().AddChild(newBlock);
            blockStack.Push(newBlock);
        }

        private ContentBlock CreateContentBlock(SyntaxBlock syntaxBlock)
        {
            return syntaxBlock.PatternType switch
            {
                PatternType.CollectionStart when syntaxBlock.MatchedStartNode.Formatter.Equals("dyntable", StringComparison.InvariantCultureIgnoreCase) => new DynamicTableBlock(syntaxBlock.MatchedStartNode.Variable, m_variableReplacer),
                PatternType.CollectionStart => new LoopBlock(syntaxBlock.MatchedStartNode.Variable, m_variableReplacer),
                PatternType.Condition => new ConditionalBlock(syntaxBlock.MatchedStartNode.Condition, m_variableReplacer, m_scriptCompiler),
                _ => new ContentBlock(m_variableReplacer)
            };
        }

        private void ExtractBlockContent(SyntaxBlock block)
        {
            var commonParent = block.ParentNode;

            OpenXmlElement anchorElement;
            block.ContentBlock = CreateContentBlock(block);
            List<OpenXmlElement> blockContent = new();

            if (commonParent is TableRow)
            {
                anchorElement = commonParent.InsertBeforeSelf(new TableRow());
                blockContent.Add(commonParent);
            }
            else if (commonParent is Table)
            {
                var firstRow = commonParent.Elements<TableRow>().First(block.StartTextNode.IsChildOf);
                var lastRow = commonParent.Elements<TableRow>().First(block.EndTextNode.IsChildOf);
                anchorElement = firstRow.InsertBeforeSelf(new TableRow());
                blockContent.Add(firstRow);
                blockContent.AddRange(commonParent.ChildsBetween(firstRow, lastRow));
                blockContent.Add(lastRow);
            }
            else
            {
                // find childs of common parent that contains start and end text
                var startChildOfCommonParent = commonParent.ChildElements.Single(c => c == block.StartTextNode || c.Descendants<Text>().Any(d => d == block.StartTextNode));
                var endChildOfCommonParent = commonParent.ChildElements.Single(c => c == block.EndTextNode || c.Descendants<Text>().Any(d => d == block.EndTextNode));
                var startSplit = startChildOfCommonParent.SplitAfterElement(block.StartTextNode);
                anchorElement = startSplit.First();
                if (startChildOfCommonParent == endChildOfCommonParent)
                {
                    blockContent.AddRange(commonParent.ChildsBetween(startSplit.First(), endChildOfCommonParent).ToList());
                }
                else
                {
                    var endSplit = endChildOfCommonParent.SplitBeforeElement(block.EndTextNode);
                    blockContent.AddRange(commonParent.ChildsBetween(anchorElement, endSplit.Last()).ToList());
                }
            }

            foreach (var contentElement in blockContent)
            {
                contentElement.Remove();
            }
            block.ContentBlock.SetContent(InsertionPoint.CreateForElement(anchorElement, block.StartTextNode.Text), blockContent);
        }

        private class SyntaxBlock
        {
            private readonly List<SyntaxBlock> m_childs = new();

            public void AddChild(SyntaxBlock block)
            {
                m_childs.Add(block);
            }

            public IReadOnlyCollection<SyntaxBlock> ChildBlocks => m_childs;

            public PatternType PatternType
            {
                get;
                set;
            }

            public Text StartTextNode
            {
                get;
                set;
            }

            public Text EndTextNode
            {
                get;
                set;
            }

            public PatternMatch MatchedStartNode
            {
                get;
                set;
            }

            public PatternMatch MatchedEndNode
            {
                get;
                set;
            }

            public OpenXmlElement ParentNode { get; set; }
            public ContentBlock ContentBlock { get; set; }

            public void AssignBlockChildren()
            {
                foreach (var child in m_childs)
                {
                    ContentBlock.AddInnerBlock(child.ContentBlock);
                    child.AssignBlockChildren();
                }
            }

            public override string ToString()
            {
                return $"{PatternType}_{StartTextNode.Text}";
            }

            public IEnumerable<SyntaxBlock> TraverseUp()
            {
                foreach (var child in m_childs)
                {
                    foreach (var item in child.TraverseUp())
                    {
                        yield return item;
                    }
                }
                yield return this;
            }
        }

    }
}
