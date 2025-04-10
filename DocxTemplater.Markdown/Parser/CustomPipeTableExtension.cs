using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Parsers.Inlines;
using Markdig.Renderers;

namespace DocxTemplater.Markdown.Parser
{
    /* This file is a modified copy from markdig project.
// required unit the feature is released in markdig
// https://github.com/xoofx/markdig/pull/863
*/

    /// <summary>
    /// Extension that allows to use pipe tables.
    /// </summary>
    /// <seealso cref="IMarkdownExtension" />
    public class CustomPipeTableExtension : IMarkdownExtension
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CustomPipeTableExtension"/> class.
        /// </summary>
        /// <param name="options">The options.</param>
        public CustomPipeTableExtension(PipeTableOptions options = null)
        {
            Options = options ?? new PipeTableOptions();
        }

        /// <summary>
        /// Gets the options.
        /// </summary>
        public PipeTableOptions Options { get; }

        public void Setup(MarkdownPipelineBuilder pipeline)
        {
            // Pipe tables require precise source location
            pipeline.PreciseSourceLocation = true;
            if (!pipeline.BlockParsers.Contains<PipeTableBlockParser>())
            {
                pipeline.BlockParsers.Insert(0, new PipeTableBlockParser());
            }
            var lineBreakParser = pipeline.InlineParsers.FindExact<LineBreakInlineParser>();
            if (!pipeline.InlineParsers.Contains<CustomPipeTableParser>())
            {
                pipeline.InlineParsers.InsertBefore<EmphasisInlineParser>(new CustomPipeTableParser(lineBreakParser!, Options));
            }
        }

        public void Setup(MarkdownPipeline pipeline, IMarkdownRenderer renderer)
        {
            if (renderer is HtmlRenderer htmlRenderer && !htmlRenderer.ObjectRenderers.Contains<HtmlTableRenderer>())
            {
                htmlRenderer.ObjectRenderers.Add(new HtmlTableRenderer());
            }
        }
    }
}