﻿
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Markdown.Renderer;
using DocxTemplater.Markdown.Renderer.Inlines;
using Markdig.Helpers;
using Markdig.Renderers;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using System;
using System.Collections.Generic;

namespace DocxTemplater.Markdown
{
    internal sealed class MarkdownToOpenXmlRenderer : RendererBase
    {
        private sealed record Format(bool Bold, bool Italic, string Style);

        private readonly Stack<Format> m_formatStack = new();
        private OpenXmlCompositeElement m_parentElement;
        private bool m_lastElemntWasNewLine;
        private readonly RunProperties m_targetRunProperties;

        public MarkdownToOpenXmlRenderer(
            Paragraph parentElement,
            Text target,
            MainDocumentPart mainDocumentPart,
            MarkDownFormatterConfiguration configuration)
        {
            // extract style from target run element
            m_targetRunProperties = ((Run)target.Parent).RunProperties;
            m_lastElemntWasNewLine = true;
            m_formatStack.Push(new Format(false, false, null));
            m_parentElement = parentElement;
            ObjectRenderers.Add(new LiteralInlineRenderer());
            ObjectRenderers.Add(new ParagraphRenderer());
            ObjectRenderers.Add(new LineBreakLineRenderer());
            ObjectRenderers.Add(new EmphasisInlineRenderer());
            ObjectRenderers.Add(new TableRenderer(configuration, mainDocumentPart));
            ObjectRenderers.Add(new ListRenderer(mainDocumentPart, configuration));
            ObjectRenderers.Add(new HeadingRenderer());
            ObjectRenderers.Add(new ThematicBreakRenderer());
        }

        public Paragraph CurrentParagraph => m_parentElement as Paragraph;
        public MarkdownToOpenXmlRenderer Write(ref StringSlice slice)
        {
            Write(slice.AsSpan());
            return this;
        }

        public void Write(ReadOnlySpan<char> content)
        {
            if (!content.IsEmpty)
            {
                var text = new Text(content.ToString());
                if (char.IsWhiteSpace(content[^1]) || char.IsWhiteSpace(content[0]))
                {
                    text.Space = SpaceProcessingModeValues.Preserve;
                }

                var newRun = new Run(text);
                if (m_targetRunProperties != null)
                {
                    // we merge existing styles with styles from markdown
                    newRun.RunProperties = (RunProperties)m_targetRunProperties.CloneNode(true);
                }
                var format = m_formatStack.Peek();
                if (format.Bold || format.Italic || format.Style != null)
                {
                    newRun.RunProperties ??= new RunProperties();

                    if (format.Bold && newRun.RunProperties.Bold == null)
                    {
                        newRun.RunProperties.AddChild(new Bold());
                    }
                    if (format.Italic && newRun.RunProperties.Italic == null)
                    {
                        newRun.RunProperties.AddChild(new Italic());
                    }
                    //add style
                    if (format.Style != null)
                    {
                        var runStyle = new RunStyle { Val = format.Style };
                        newRun.RunProperties.AddChild(runStyle);
                    }
                }
                m_parentElement.Append(newRun);
                m_lastElemntWasNewLine = false;
            }
        }

        public override object Render(MarkdownObject markdownObject)
        {
            Write(markdownObject);
            return null;
        }

        public void WriteLeafInline(LeafBlock leafBlock)
        {
            Inline inline = leafBlock.Inline;
            while (inline != null)
            {
                Write(inline);
                inline = inline.NextSibling;
            }
        }

        public IDisposable PushFormat(bool? bold, bool? italic)
        {
            var currentStyle = m_formatStack.Peek();
            bold ??= currentStyle.Bold;
            italic ??= currentStyle.Italic;
            return new FormatScope(m_formatStack, bold.Value, italic.Value, currentStyle.Style);
        }

        public void NewLine()
        {
            m_parentElement.Append(new Run(new Break()));
            m_lastElemntWasNewLine = true;
        }

        public void EnsureNewLine()
        {
            if (!m_lastElemntWasNewLine)
            {
                NewLine();
            }
        }

        public void ReplaceIfCurrentParagraphIsEmpty(OpenXmlCompositeElement newParagraph)
        {
            var lastParagraph = CurrentParagraph;
            AddParagraph(newParagraph);
            if (lastParagraph != null && lastParagraph.ChildElements.Count == 0)
            {
                lastParagraph.Remove();
            }
        }

        public void AddParagraph(OpenXmlCompositeElement paragraph = null)
        {
            paragraph ??= new Paragraph();
            m_parentElement = m_parentElement.InsertAfterSelf(paragraph);
            m_lastElemntWasNewLine = false;
        }

        public IDisposable PushParagraph(Paragraph paragraph)
        {
            m_lastElemntWasNewLine = true;
            return new ParagraphScope(this, paragraph);
        }

        private sealed class ParagraphScope : IDisposable
        {
            private readonly MarkdownToOpenXmlRenderer m_renderer;
            private readonly OpenXmlCompositeElement m_previousParagraph;

            public ParagraphScope(MarkdownToOpenXmlRenderer renderer, Paragraph element)
            {
                m_renderer = renderer;
                m_previousParagraph = m_renderer.m_parentElement;
                m_renderer.m_parentElement = element;
            }

            public void Dispose()
            {
                m_renderer.m_parentElement = m_previousParagraph;
            }
        }

        private sealed class FormatScope : IDisposable
        {
            private readonly Stack<Format> m_formatStack;
            public FormatScope(Stack<Format> formatStack, bool bold, bool italic, string style)
            {
                m_formatStack = formatStack;
                m_formatStack.Push(new Format(bold, italic, style));
            }

            public void Dispose()
            {
                m_formatStack.Pop();
            }
        }
    }

}
