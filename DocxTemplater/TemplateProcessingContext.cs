using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.Formatter;

namespace DocxTemplater
{
    public interface ITemplateProcessingContext
    {
        ProcessSettings ProcessSettings { get; }
        MainDocumentPart MainDocumentPart { get; }
        IScriptCompiler ScriptCompiler { get; }
        IModelLookup ModelLookup { get; }
        IVariableReplacer VariableReplacer { get; }
    }

    public interface ITemplateProcessingContextAccess : ITemplateProcessingContext
    {
        void Initialize(MainDocumentPart mainDocumentPart);
    }

    internal class TemplateProcessingContext : ITemplateProcessingContextAccess
    {
        public ProcessSettings ProcessSettings { get; }
        public MainDocumentPart MainDocumentPart { get; set; }
        public IScriptCompiler ScriptCompiler { get; }
        public IModelLookup ModelLookup { get; }
        public IVariableReplacer VariableReplacer { get; }
        public void Initialize(MainDocumentPart mainDocumentPart)
        {
            MainDocumentPart = mainDocumentPart;
        }

        public TemplateProcessingContext(ProcessSettings processSettings, IModelLookup modelLookup, IVariableReplacer variableReplacer, IScriptCompiler scriptCompiler)
        {
            ProcessSettings = processSettings;
            ModelLookup = modelLookup;
            VariableReplacer = variableReplacer;
            ScriptCompiler = scriptCompiler;
        }
    }
}
