using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class ParagraphRenderer : OpenXmlObjectRenderer<ParagraphBlock>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, ParagraphBlock obj)
        {
            renderer.WriteLeafInline(obj);
            if (!renderer.IsLastInContainer)
            {
                var paragraph = new Paragraph();
                if (obj.Parent is MarkdownDocument)
                {

                    // if a new paragraph is created, we need to copy the paragraph properties from the template
                    paragraph = new Paragraph();
                    var paragraphProperties = renderer.GetTemplateParagraphProperties();
                    if (paragraphProperties != null)
                    {
                        paragraph.Append(paragraphProperties);
                    }

                }

                renderer.AddParagraph(paragraph);
            }
        }
    }
}