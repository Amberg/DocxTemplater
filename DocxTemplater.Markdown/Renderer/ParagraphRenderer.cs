using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class ParagraphRenderer : OpenXmlObjectRenderer<ParagraphBlock>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, ParagraphBlock obj)
        {
            renderer.WriteLeafInline(obj);
            if (renderer.ExplicitParagraph)
            {
                renderer.ExplicitParagraph = false;
                return;
            }
            renderer.AddParagraph();
        }
    }
}