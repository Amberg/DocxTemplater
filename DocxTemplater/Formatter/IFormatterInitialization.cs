using DocumentFormat.OpenXml.Packaging;

namespace DocxTemplater.Formatter
{

    /// <summary>
    /// Implement this interface on a <see cref="IFormatter"/> If the formatter needs to be initialized with the <see cref="IModelLookup"/> and <see cref="ProcessSettings"/>
    /// </summary>
    public interface IFormatterInitialization
    {
        void Initialize(IModelLookup modelLookup, IScriptCompiler scriptCompiler, IVariableReplacer variableReplacer, ProcessSettings processSettings, MainDocumentPart mainDocumentPart);
    }
}
