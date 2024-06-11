using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    internal sealed class ThematicBreakRenderer : OpenXmlObjectRenderer<ThematicBreakBlock>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, ThematicBreakBlock obj)
        {
            renderer.AddParagraph(CreateParagraphWithBorder());
            renderer.AddParagraph();
        }

        private static Paragraph CreateParagraphWithBorder()
        {
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties();
            var borders = CreateBorders();
            paragraphProperties.Append(borders);
            paragraph.Append(paragraphProperties);
            return paragraph;
        }

        private static ParagraphBorders CreateBorders()
        {
            var borders = new ParagraphBorders();
            var bottomBorder = new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Color = "auto", Size = 6 };
            borders.Append(bottomBorder);
            return borders;
        }
    }
}
