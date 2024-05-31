using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocxTemplater.Formatter;

namespace DocxTemplater.Blocks
{
    internal class ContentBlock
    {
        protected InsertionPoint m_insertionPoint;
        protected IReadOnlyCollection<OpenXmlElement> m_content;
        protected readonly List<ContentBlock> m_childBlocks;
        protected readonly VariableReplacer m_variableReplacer;

        public ContentBlock(VariableReplacer variableReplacer)
        {
            m_content = new List<OpenXmlElement>();
            m_childBlocks = new List<ContentBlock>();
            m_variableReplacer = variableReplacer;
        }

        public IReadOnlyCollection<ContentBlock> ChildBlocks => m_childBlocks;

        public virtual void Expand(ModelLookup models, OpenXmlElement parentNode, bool insertBeforeInsertionPoint = false)
        {
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned, insertBeforeInsertionPoint);
            m_variableReplacer.ReplaceVariables(cloned);
            ExpandChildBlocks(models, parentNode);
        }

        protected void ExpandChildBlocks(ModelLookup models, OpenXmlElement parentNode)
        {
            foreach (var child in m_childBlocks)
            {
                child.Expand(models, parentNode);
            }
        }

        protected void InsertContent(OpenXmlElement parentNode, IEnumerable<OpenXmlElement> paragraphs, bool insertBeforeInsertionPoint = false)
        {
            var element = m_insertionPoint.GetElement(parentNode) ?? throw new OpenXmlTemplateException($"Insertion point {m_insertionPoint.Id} not found");
            if (insertBeforeInsertionPoint)
            {
                element.InsertBeforeSelf(paragraphs);
            }
            else
            {
                element.InsertAfterSelf(paragraphs);
            }
        }

        public void RemoveAnchor(OpenXmlElement parentNode)
        {
            var element = m_insertionPoint.GetElement(parentNode) ?? throw new OpenXmlTemplateException($"Insertion point {m_insertionPoint.Id} not found");
            element.Remove();
        }

        public override string ToString()
        {
            return m_insertionPoint?.Id ?? GetType().Name;
        }

        public virtual void SetContent(InsertionPoint insertionPoint, IReadOnlyCollection<OpenXmlElement> blockContent)
        {
            m_insertionPoint ??= insertionPoint;
            m_content = blockContent;
        }

        public virtual void AddInnerBlock(ContentBlock block)
        {
            m_childBlocks.Add(block);
        }

        public void Print(int i)
        {
            Console.WriteLine(new string('-', i + 1) + ToString());
            foreach (var content in m_content)
            {
                Console.WriteLine(content.ToPrettyPrintXml());
            }
            foreach (var block in m_childBlocks)
            {
                block.Print(i + 2);
            }
        }
    }
}