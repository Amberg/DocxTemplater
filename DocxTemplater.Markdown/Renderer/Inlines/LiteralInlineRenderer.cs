using Markdig.Syntax.Inlines;

namespace DocxTemplater.Markdown.Renderer.Inlines
{
    internal sealed class LiteralInlineRenderer : OpenXmlObjectRenderer<LiteralInline>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, LiteralInline obj)
        {
            renderer.Write(ref obj.Content);
        }
    }
}