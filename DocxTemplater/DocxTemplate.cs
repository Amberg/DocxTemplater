using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocxTemplater.Formatter;
using System;
using System.IO;
using System.Linq;
using DocxTemplater.Extensions.Charts;

namespace DocxTemplater
{
    public sealed class DocxTemplate : TemplateProcessor, IDisposable
    {
        private readonly Stream m_stream;
        private readonly WordprocessingDocument m_wpDocument;

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
        : base(new TemplateProcessingContext(settings, modelLookup, new VariableReplacer(modelLookup, settings), new ScriptCompiler(modelLookup, settings)))
        {
            ArgumentNullException.ThrowIfNull(docXStream);
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
                ProcessNode(header.Header);
            }
            ProcessNode(m_wpDocument.MainDocumentPart.RootElement);
            foreach (var footer in m_wpDocument.MainDocumentPart.FooterParts.ToList())
            {
                ProcessNode(footer.Footer);
            }
            Context.VariableReplacer.WriteErrorMessages(m_wpDocument.MainDocumentPart.RootElement);
            m_wpDocument.Save();
            m_stream.Position = 0;
            return m_stream;
        }


        public void Dispose()
        {
            m_stream?.Dispose();
            m_wpDocument?.Dispose();
        }
    }
}
