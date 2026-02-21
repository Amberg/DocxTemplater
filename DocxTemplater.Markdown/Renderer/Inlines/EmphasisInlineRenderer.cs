using Markdig.Syntax.Inlines;

namespace DocxTemplater.Markdown.Renderer.Inlines
{
    internal sealed class EmphasisInlineRenderer : OpenXmlObjectRenderer<EmphasisInline>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, EmphasisInline obj)
        {
            bool? italic = null;
            bool? bold = null;
            bool? subscript = null;
            bool? superscript = null;
            bool? strikethrough = null;

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

            if (obj.DelimiterChar is '~')
            {
                if (obj.DelimiterCount == 1)
                {
                    subscript = true;
                }
                else if (obj.DelimiterCount == 2)
                {
                    strikethrough = true;
                }
            }

            if (obj.DelimiterChar is '^')
            {
                if (obj.DelimiterCount == 1)
                {
                    superscript = true;
                }
            }

            using var format = renderer.PushFormat(bold, italic, strikethrough, subscript, superscript);
            renderer.WriteChildren(obj);
        }
    }
}