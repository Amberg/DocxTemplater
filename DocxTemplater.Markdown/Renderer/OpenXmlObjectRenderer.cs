using Markdig.Renderers;
using Markdig.Syntax;

namespace DocxTemplater.Markdown.Renderer
{
    internal abstract class OpenXmlObjectRenderer<TObject> : MarkdownObjectRenderer<MarkdownToOpenXmlRenderer, TObject> where TObject : MarkdownObject
    {
    }
}