using DocumentFormat.OpenXml.Packaging;
using DocxTemplater.Formatter;

namespace DocxTemplater
{
    public class TemplateProcessingContext
    {
        public ProcessSettings ProcessSettings { get; }
        public MainDocumentPart MainDocumentPart { get; set; }
        public IScriptCompiler ScriptCompiler { get; }
        public IModelLookup ModelLookup { get; }
        public IVariableReplacer VariableReplacer { get; }

        public TemplateProcessingContext(ProcessSettings processSettings, IModelLookup modelLookup, IVariableReplacer variableReplacer, IScriptCompiler scriptCompiler)
        {
            ProcessSettings = processSettings;
            ModelLookup = modelLookup;
            VariableReplacer = variableReplacer;
            ScriptCompiler = scriptCompiler;
        }
    }
}
