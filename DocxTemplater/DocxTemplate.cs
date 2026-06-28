using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxTemplater.Blocks;
using DocxTemplater.Formatter;
using DocxTemplater.Schema;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocxTemplater.Extensions.Charts;

namespace DocxTemplater
{
    public sealed class DocxTemplate : TemplateProcessor, IDisposable
    {
        private readonly Stream m_stream;
        private readonly WordprocessingDocument m_wpDocument;

        // Block trees built by GetTemplateSchema, keyed by the node they were built from (header /
        // root / footer). Process reuses these instead of rebuilding the (now mutated) tree.
        private readonly Dictionary<OpenXmlCompositeElement, IReadOnlyCollection<ContentBlock>> m_prebuiltBlocks = new();
        private TemplateSchema m_schema;

        private static readonly FileFormatVersions TargetMinimumVersion = FileFormatVersions.Office2010;

        public DocxTemplate(Stream docXStream)
        : this(docXStream, ProcessSettings.Default)
        {
        }

        public DocxTemplate(Stream docXStream, ProcessSettings settings)
            : this(docXStream, settings, new ModelLookup())
        {
        }

        private DocxTemplate(Stream docXStream, ProcessSettings settings, ModelLookup modelLookup)
        : base(CreateContext(settings, modelLookup))
        {
            if (docXStream == null)
            {
                throw new ArgumentNullException(nameof(docXStream));
            }
            m_stream = new MemoryStream();
            docXStream.CopyTo(m_stream);
            m_stream.Position = 0;
            // Add the MarkupCompatibilityProcessSettings
            OpenSettings openSettings = new()
            {
                MarkupCompatibilityProcessSettings =
                new MarkupCompatibilityProcessSettings(
                    MarkupCompatibilityProcessMode.ProcessAllParts,
                    TargetMinimumVersion)
            };
            m_wpDocument = WordprocessingDocument.Open(m_stream, true, openSettings);
            Context.Initialize(m_wpDocument.MainDocumentPart);
            Processed = false;

            RegisterFormatter(new SubTemplateFormatter());
            RegisterExtension(new ChartProcessor());
        }

        private static TemplateProcessingContext CreateContext(ProcessSettings settings, ModelLookup modelLookup)
        {
            return new TemplateProcessingContext(settings, modelLookup, new VariableReplacer(modelLookup, settings), new ScriptCompiler(modelLookup, settings));
        }

        public static DocxTemplate Open(string pathToTemplate, ProcessSettings settings = null)
        {
            settings ??= ProcessSettings.Default;
            using var fileStream = new FileStream(pathToTemplate, FileMode.Open, FileAccess.Read);
            return new DocxTemplate(fileStream, settings);
        }

        public bool Processed { get; private set; }

        public ProcessSettings Settings => Context.ProcessSettings;

        public void Save(string targetPath)
        {
            using var fileStream = new FileStream(targetPath, FileMode.Create, FileAccess.Write);
            Save(fileStream);
        }

        public void Save(Stream target)
        {
            Process().CopyTo(target);
        }

        public Stream AsStream()
        {
            return Process();
        }

        public void Validate()
        {
            var v = new OpenXmlValidator(TargetMinimumVersion);
            var errs = v.Validate(m_wpDocument);
            if (errs.Any())
            {
                var sb = new System.Text.StringBuilder();
                foreach (var err in errs)
                {
                    sb.AppendLine($"{err.Description} Node: {err.Node} RelatedNode: {err.RelatedNode}");
                }
                throw new OpenXmlTemplateException($"Validation failed - {sb}");
            }
        }

        public Stream Process()
        {
            if (m_wpDocument.MainDocumentPart == null || Processed)
            {
                m_stream.Position = 0;
                return m_stream;
            }
            Processed = true;
            foreach (var header in m_wpDocument.MainDocumentPart.HeaderParts.ToList())
            {
                RenderOrProcessNode(header.Header);
            }
            RenderOrProcessNode(m_wpDocument.MainDocumentPart.RootElement);
            foreach (var footer in m_wpDocument.MainDocumentPart.FooterParts.ToList())
            {
                RenderOrProcessNode(footer.Footer);
            }
            Context.VariableReplacer.WriteErrorMessages(m_wpDocument.MainDocumentPart.RootElement);
            m_wpDocument.Save();
            m_stream.Position = 0;
            return m_stream;
        }

        private void RenderOrProcessNode(OpenXmlCompositeElement node)
        {
            if (node == null)
            {
                return;
            }
            // If GetTemplateSchema already built the block tree for this node, render from it instead
            // of rebuilding (the tree was extracted from the node, so a rebuild would find nothing /
            // corrupt the document). Otherwise run the full pipeline as usual.
            if (m_prebuiltBlocks.TryGetValue(node, out var blocks))
            {
                RenderNode(node, blocks);
            }
            else
            {
                ProcessNode(node);
            }
        }


        /// <summary>
        /// Statically analyzes the template (without rendering it) and returns the structural schema
        /// of the variables, collections, and nested objects the template references. Use this to
        /// validate a caller's model against the template's expectations before rendering, or to
        /// generate a skeleton object for callers to fill in.
        /// </summary>
        /// <remarks>
        /// This method builds the internal block tree (mutates the XML) but does not render. The
        /// built tree is cached, so a subsequent call to <see cref="Process"/> renders from it
        /// instead of rebuilding; calling <c>GetTemplateSchema</c> and then <c>Process</c> on the same
        /// instance is supported (and the result is itself cached, so repeated calls are cheap).
        /// It must be called before <see cref="Process"/>, not after. The schema represents the union
        /// of all branches (both <c>if</c> and <c>else</c>, all <c>case</c>s, every loop body) -
        /// callers must be prepared to bind anything the template could reach at runtime.
        /// See <see cref="TemplateSchema"/> for known limitations.
        /// </remarks>
        public TemplateSchema GetTemplateSchema()
        {
            if (m_schema != null)
            {
                return m_schema;
            }
            if (Processed)
            {
                throw new OpenXmlTemplateException($"{nameof(GetTemplateSchema)} must be called before {nameof(Process)}.");
            }
            var builder = new SchemaBuilder();
            if (m_wpDocument.MainDocumentPart != null)
            {
                foreach (var header in m_wpDocument.MainDocumentPart.HeaderParts.ToList())
                {
                    CollectFromPart(builder, header.Header);
                }
                CollectFromPart(builder, m_wpDocument.MainDocumentPart.RootElement);
                foreach (var footer in m_wpDocument.MainDocumentPart.FooterParts.ToList())
                {
                    CollectFromPart(builder, footer.Footer);
                }
            }
            m_schema = builder.Build();
            return m_schema;
        }

        private void CollectFromPart(SchemaBuilder builder, OpenXmlCompositeElement rootElement)
        {
            if (rootElement == null)
            {
                return;
            }
            var rootBlocks = BuildBlockTree(rootElement);
            // Cache the built tree so Process renders from it instead of rebuilding the mutated node.
            m_prebuiltBlocks[rootElement] = rootBlocks;
            // Root-level variables outside any block remain as marked Text nodes in the rootElement
            // after the block-tree pipeline extracts each block's content. Pick those up directly.
            ContentBlock.CollectMarkedVariables([rootElement], builder);
            foreach (var block in rootBlocks)
            {
                block.CollectSchema(builder);
            }
        }

        public void Dispose()
        {
            m_wpDocument?.Dispose();
            m_stream?.Dispose();
        }
    }
}

