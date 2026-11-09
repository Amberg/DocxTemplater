using Markdig.Syntax.Inlines;

namespace DocxTemplater.Markdown.Renderer.Inlines
{
    /// <summary>Inline HTML and HTML entities are written as text, so <c>&lt;</c> / <c>&amp;lt;</c> in a value stay visible.</summary>
    internal sealed class HtmlInlineRenderer : OpenXmlObjectRenderer<HtmlInline>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, HtmlInline obj)
        {
            renderer.Write(obj.Tag);
        }
    }

    internal sealed class HtmlEntityInlineRenderer : OpenXmlObjectRenderer<HtmlEntityInline>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, HtmlEntityInline obj)
        {
            renderer.Write(obj.Transcoded.AsSpan());
        }
    }
}
