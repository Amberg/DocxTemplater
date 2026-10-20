namespace DocxTemplater.Text
{
    /// <summary>
    /// What a <see cref="ITextFormatter"/> can access while a <see cref="TextTemplate"/> is rendered.
    /// </summary>
    public interface ITextTemplateContext
    {
        TextTemplateSettings Settings { get; }

        IModelLookup ModelLookup { get; }

        IScriptCompiler ScriptCompiler { get; }
    }

    internal sealed class TextTemplateContext : ITextTemplateContext
    {
        public TextTemplateContext(TextTemplateSettings settings, IModelLookup modelLookup, IScriptCompiler scriptCompiler)
        {
            Settings = settings;
            ModelLookup = modelLookup;
            ScriptCompiler = scriptCompiler;
        }

        public TextTemplateSettings Settings { get; }

        public IModelLookup ModelLookup { get; }

        public IScriptCompiler ScriptCompiler { get; }
    }
}
