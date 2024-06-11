using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class HeadingRenderer : OpenXmlObjectRenderer<HeadingBlock>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, HeadingBlock heading)
        {
            var headingParagraph = new Paragraph();
            // add heading style
            var headingStyle = new ParagraphStyleId() { Val = $"Heading{heading.Level}" };
            var paragraphProps = new ParagraphProperties();
            paragraphProps.Append(headingStyle);
            headingParagraph.ParagraphProperties = paragraphProps;
            renderer.AddParagraph(headingParagraph);
            renderer.WriteLeafInline(heading);
            renderer.AddParagraph();
        }
    }
}
