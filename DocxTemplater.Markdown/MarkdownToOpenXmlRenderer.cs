
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
using System.Text;
using DocxTemplater.ImageBase;

namespace DocxTemplater.Markdown
{
    internal sealed class MarkdownToOpenXmlRenderer : RendererBase
    {
        private sealed record Format(bool Bold, bool Italic, string Style, bool Strike);

        private readonly Stack<Format> m_formatStack = new();
        private readonly RunProperties m_targetRunProperties;
        private readonly Paragraph m_containingParagraphFromTemplate;
#if DEBUG
        private readonly MarkdownDebugOutput m_debugOutput;
#endif

        public MarkdownToOpenXmlRenderer(Paragraph parentElement,
            Text target,
            MainDocumentPart mainDocumentPart,
            MarkDownFormatterConfiguration configuration,
            IImageService imageService)
        {
            // extract style from target run element
            m_targetRunProperties = ((Run)target.Parent).RunProperties;
            m_formatStack.Push(new Format(false, false, null, false));
            m_containingParagraphFromTemplate = parentElement;
            CurrentParagraph = parentElement;
            ObjectRenderers.Add(new LiteralInlineRenderer());
            ObjectRenderers.Add(new ParagraphRenderer());
            ObjectRenderers.Add(new LineBreakLineRenderer());
            ObjectRenderers.Add(new EmphasisInlineRenderer());
            ObjectRenderers.Add(new TableRenderer(configuration, mainDocumentPart));
            ObjectRenderers.Add(new ListRenderer(mainDocumentPart, configuration));
            ObjectRenderers.Add(new HeadingRenderer());
            ObjectRenderers.Add(new ThematicBreakRenderer());
            ObjectRenderers.Add(new HtmlBlockRenderer());
            if (imageService != null)
            {
                ObjectRenderers.Add(new ImageInlineRenderer(mainDocumentPart, imageService));
            }


#if DEBUG
            m_debugOutput = new MarkdownDebugOutput(this);
#endif

        }

        public Paragraph CurrentParagraph { get; private set; }

        public bool CurrentParagraphWasCreatedByMarkdown => CurrentParagraph != m_containingParagraphFromTemplate;

#if DEBUG
        public string MarkdownStructureAsString => m_debugOutput?.ToString();
#endif

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
                if (format.Bold || format.Italic || format.Strike || format.Style != null)
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

                    if (format.Strike && newRun.RunProperties.Strike == null)
                    {
                        newRun.RunProperties.AddChild(new Strike());
                    }

                    //add style
                    if (format.Style != null)
                    {
                        var runStyle = new RunStyle { Val = format.Style };
                        newRun.RunProperties.AddChild(runStyle);
                    }
                }

                CurrentParagraph.Append(newRun);
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

        public IDisposable PushFormat(bool? bold, bool? italic, bool? strike)
        {
            var currentStyle = m_formatStack.Peek();
            bold ??= currentStyle.Bold;
            italic ??= currentStyle.Italic;
            strike ??= currentStyle.Strike;
            return new FormatScope(m_formatStack, bold.Value, italic.Value, strike.Value, currentStyle.Style);
        }

        public void NewLine()
        {
            CurrentParagraph.Append(new Run(new Break()));
        }

        public void ReplaceIfCurrentParagraphIsEmpty(Paragraph newParagraph)
        {
            var lastParagraph = CurrentParagraph;
            AddParagraph(newParagraph);
            if (lastParagraph != null && lastParagraph != m_containingParagraphFromTemplate &&
                lastParagraph.HasOnlyPropertyChildren())
            {
                lastParagraph.Remove();
            }
        }

        public void AddParagraph(Paragraph paragraph = null)
        {
            paragraph ??= new Paragraph();
            CurrentParagraph = CurrentParagraph.InsertAfterSelf(paragraph);
        }

        public void InsertNonParagraphContainer(OpenXmlCompositeElement compositeElement)
        {
            CurrentParagraph.InsertAfterSelf(compositeElement);
            CurrentParagraph = new Paragraph();
            compositeElement.InsertAfterSelf(CurrentParagraph);
        }

        public IDisposable PushParagraph(Paragraph paragraph)
        {
            return new ParagraphScope(this, paragraph);
        }

        private sealed class ParagraphScope : IDisposable
        {
            private readonly MarkdownToOpenXmlRenderer m_renderer;
            private readonly Paragraph m_previousParagraph;

            public ParagraphScope(MarkdownToOpenXmlRenderer renderer, Paragraph element)
            {
                m_renderer = renderer;
                m_previousParagraph = m_renderer.CurrentParagraph;
                m_renderer.CurrentParagraph = element;
            }

            public void Dispose()
            {
                m_renderer.CurrentParagraph = m_previousParagraph;
            }
        }

        private sealed class FormatScope : IDisposable
        {
            private readonly Stack<Format> m_formatStack;

            public FormatScope(Stack<Format> formatStack, bool bold, bool italic, bool strike, string style)
            {
                m_formatStack = formatStack;
                m_formatStack.Push(new Format(bold, italic, style, strike));
            }

            public void Dispose()
            {
                m_formatStack.Pop();
            }
        }

        private sealed class MarkdownDebugOutput
        {
            private readonly RendererBase m_renderer;
            private int m_indent;
            private readonly StringBuilder m_buffer = new();

            public MarkdownDebugOutput(RendererBase renderer)
            {
                m_renderer = renderer;
                m_buffer.AppendLine("---------- Markdown Document ---------");
                Register();
            }

            public override string ToString()
            {
                return m_buffer.ToString();
            }

            private void Register()
            {
                m_renderer.ObjectWriteBefore += (_, markdownObject) =>
                {
                    m_indent++;
                    m_buffer.Append(new string(' ', m_indent * 2));
                    m_buffer.Append(markdownObject.GetType().Name);
                    if (markdownObject is LiteralInline)
                    {
                        m_buffer.Append($" : '{markdownObject.ToString()}'");
                    }

                    m_buffer.AppendLine();
                };

                m_renderer.ObjectWriteAfter += (renderer, markdownObject) =>
                {
                    m_indent--;
                    if (markdownObject is MarkdownDocument)
                    {
                        m_buffer.AppendLine("---------- END Markdown Document ---------");
                    }
                };
            }
        }

        public ParagraphProperties GetTemplateParagraphProperties()
        {
            return (ParagraphProperties)m_containingParagraphFromTemplate?.ParagraphProperties?.CloneNode(true);
        }
    }
}
