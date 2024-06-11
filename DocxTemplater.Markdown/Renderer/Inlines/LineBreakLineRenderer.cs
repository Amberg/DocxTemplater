using Markdig.Syntax.Inlines;

namespace DocxTemplater.Markdown.Renderer.Inlines
{
    internal sealed class LineBreakLineRenderer : OpenXmlObjectRenderer<LineBreakInline>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, LineBreakInline obj)
        {
            if (renderer.IsLastInContainer)
            {
                return;
            }
            renderer.NewLine();
        }
    }
}
