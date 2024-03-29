﻿using System;
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

        public ContentBlock(VariableReplacer variableReplacer, ContentBlock rootBlock = null, InsertionPoint insertionPoint = null)
        {
            m_insertionPoint = insertionPoint;
            RootBlock = rootBlock ?? this;
            m_content = new List<OpenXmlElement>();
            m_childBlocks = new List<ContentBlock>();
            m_variableReplacer = variableReplacer;
        }

        public IReadOnlyCollection<ContentBlock> ChildBlocks => m_childBlocks;

        public ContentBlock RootBlock
        {
            get;
        }

        public virtual void Expand(ModelLookup models, OpenXmlElement parentNode)
        {
            var cloned = m_content.Select(x => x.CloneNode(true)).ToList();
            InsertContent(parentNode, cloned);
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

        protected void InsertContent(OpenXmlElement parentNode, IEnumerable<OpenXmlElement> paragraphs)
        {
            var element = m_insertionPoint.GetElement(parentNode);
            if (element == null)
            {
                Console.WriteLine(parentNode.ToPrettyPrintXml());
                throw new OpenXmlTemplateException($"Insertion point {m_insertionPoint.Id} not found");
            }
            element.InsertAfterSelf(paragraphs);
        }

        public override string ToString()
        {
            return m_insertionPoint?.Id ?? "RootBlock";
        }

        public virtual void SetContent(InsertionPoint insertionPoint, IReadOnlyCollection<OpenXmlElement> blockContent)
        {
            m_insertionPoint ??= insertionPoint;
            m_content = blockContent;
        }

        public void AddInnerBlock(ContentBlock block)
        {
            m_childBlocks.Add(block);
        }
    }
}