using Markdig.Syntax.Inlines;

namespace DocxTemplater.Markdown.Renderer.Inlines
{
    internal sealed class EmphasisInlineRenderer : OpenXmlObjectRenderer<EmphasisInline>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, EmphasisInline obj)
        {
            bool? italic = null;
            bool? bold = null;
            if (obj.DelimiterChar is '_' or '*')
            {
                if (obj.DelimiterCount == 1)
                {
                    italic = true;
                }
                else if (obj.DelimiterCount == 2)
                {
                    bold = true;
                }
            }
            using var format = renderer.PushFormat(bold, italic);
            renderer.WriteChildren(obj);
        }
    }
}