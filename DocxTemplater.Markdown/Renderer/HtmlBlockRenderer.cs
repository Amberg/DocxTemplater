using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Markdig.Helpers;
using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    class HtmlBlockRenderer : OpenXmlObjectRenderer<HtmlBlock>
    {
        protected override void Write(MarkdownToOpenXmlRenderer renderer, HtmlBlock obj)
        {
            foreach (StringLine line in obj.Lines)
            {
                if (line.Slice.Text == "<br>")
                {
                    renderer.NewLine();
                }
            }
        }
    }

}
