using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.Formatter;
using System;
using System.Collections.Generic;
using System.Linq;
using DocxTemplater.Extensions;

namespace DocxTemplater
{
    public interface ITemplateProcessingContext
    {
        ProcessSettings ProcessSettings { get; }
        MainDocumentPart MainDocumentPart { get; }
        IScriptCompiler ScriptCompiler { get; }
        IModelLookup ModelLookup { get; }
        IVariableReplacer VariableReplacer { get; }
        IReadOnlyCollection<ITemplateProcessorExtension> Extensions { get; }
    }

    public interface ITemplateProcessingContextAccess : ITemplateProcessingContext
    {
        void Initialize(MainDocumentPart mainDocumentPart);

        void RegisterExtension(ITemplateProcessorExtension extension);
    }

    internal class TemplateProcessingContext : ITemplateProcessingContextAccess
    {
        private readonly List<ITemplateProcessorExtension> m_extensions;

        public TemplateProcessingContext(ProcessSettings processSettings, IModelLookup modelLookup, IVariableReplacer variableReplacer, IScriptCompiler scriptCompiler)
        {
            ProcessSettings = processSettings;
            ModelLookup = modelLookup;
            VariableReplacer = variableReplacer;
            ScriptCompiler = scriptCompiler;
            m_extensions = new List<ITemplateProcessorExtension>();
        }

        public ProcessSettings ProcessSettings { get; }
        public MainDocumentPart MainDocumentPart { get; set; }
        public IScriptCompiler ScriptCompiler { get; }
        public IModelLookup ModelLookup { get; }
        public IVariableReplacer VariableReplacer { get; }

        public IReadOnlyCollection<ITemplateProcessorExtension> Extensions => m_extensions;

        public void Initialize(MainDocumentPart mainDocumentPart)
        {
            MainDocumentPart = mainDocumentPart;
        }

        public void RegisterExtension(ITemplateProcessorExtension extension)
        {
            if (m_extensions.All(x => x.GetType() != extension.GetType()))
            {
                m_extensions.Add(extension);
            }
            else
            {
                throw new InvalidOperationException($"Extension of type {extension.GetType()} is already registered");
            }
            m_extensions.Add(extension);
        }
    }
}
