using System;
using System.Collections.Generic;
using System.Linq;
using DocxTemplater.Schema;

namespace DocxTemplater.Text
{
    /// <summary>
    /// A plain text or HTML template with the syntax of DocxTemplater: the same placeholders, formatters,
    /// expressions, loops, conditions and switch blocks as in a Word template, rendered to a string.
    /// </summary>
    /// <remarks>
    /// To render a template for many models - e.g. a mail per recipient - create a new instance per model with the
    /// same <see cref="TextTemplateSettings"/>: the conditions and expressions are cached in the settings, so a new
    /// instance only parses the text. The settings are copied when the instance is created. Instances with the same
    /// settings can be processed in parallel; one instance is not thread-safe. Constructs that insert Word content
    /// (<c>:img()</c>, <c>:template()</c>, <c>:md()</c>, <c>:dyntable()</c>, <c>{{:PageBreak}}</c>,
    /// <c>{{:SectionBreak}}</c>) are syntax errors.
    /// </remarks>
    public sealed class TextTemplate
    {
        /// <summary>The <see cref="TemplateSyntaxError.Part"/> of the errors of a text template.</summary>
        public const string PartName = "Text";

        private readonly string m_template;
        private readonly TextTemplateSettings m_settings;
        private readonly ModelLookup m_models;
        private ParsedTextTemplate m_parsed;
        private IReadOnlyList<TemplateSyntaxError> m_syntaxErrors;
        private List<TemplateSyntaxError> m_blockingErrors;

        public TextTemplate(string template)
            : this(template, TextTemplateSettings.Default)
        {
        }

        public TextTemplate(string template, TextTemplateSettings settings)
        {
            ArgumentNullException.ThrowIfNull(template);
            ArgumentNullException.ThrowIfNull(settings);
            m_template = template;
            m_settings = settings.Copy();
            m_models = new ModelLookup(m_settings.ProcessSettings);
        }

        /// <summary>
        /// Binds a model under a prefix: <c>{{prefix.Property}}</c>. The prefix of the first model can be omitted in the
        /// template, <c>{{Property}}</c>.
        /// </summary>
        /// <exception cref="TemplateException">A model is already bound under <paramref name="prefix"/>.</exception>
        public void BindModel(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        /// <summary>
        /// Checks the template syntax without rendering it: that all blocks are closed, closing tags match their
        /// opening tags, tags are well-formed and no construct is used that exists only in Word templates. Binding
        /// errors (unknown variables, wrong types) are not detected - they require a model. Never throws.
        /// </summary>
        /// <returns>
        /// The syntax errors and warnings, ordered by position - e.g. text that looks like a tag but is not recognized
        /// as one (<see cref="TemplateErrorCode.InvalidTag"/>, <c>{{first name}}</c>), which stays in the output as it is.
        /// Entries with <see cref="TemplateSyntaxErrorSeverity.Error"/> make <see cref="Process"/> throw, except with
        /// <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>.
        /// </returns>
        public IReadOnlyList<TemplateSyntaxError> ValidateTemplateSyntax()
        {
            Parse();
            return m_syntaxErrors;
        }

        /// <summary>
        /// Statically analyzes the template (without rendering it) and returns the structural schema of the
        /// variables, collections and nested objects it references - the union of all branches, see <see cref="TemplateSchema"/>.
        /// </summary>
        /// <exception cref="TemplateException">The template has syntax errors.</exception>
        public TemplateSchema GetTemplateSchema()
        {
            var parsed = Parse();
            ThrowIfSyntaxErrors();
            return TextTemplateAnalyzer.CollectSchema(parsed.Tree, parsed.Text.Length);
        }

        /// <summary>Renders the template with the bound models.</summary>
        /// <returns>
        /// The text and the binding errors. With <see cref="BindingErrorHandling.SkipBindingAndRemoveContent"/> the
        /// failed tags are left out of the text and only reported in <see cref="TextRenderResult.Errors"/>, so the
        /// caller can decide per render whether to use it.
        /// </returns>
        /// <exception cref="TemplateException">
        /// The template has syntax errors (<see cref="TemplateErrorCode.TemplateSyntaxErrors"/>; its argument is the list
        /// of <see cref="TemplateSyntaxError"/>s) and <see cref="TextTemplateSettings.BindingErrorHandling"/> is not
        /// <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>; or a tag could not be rendered and it is
        /// <see cref="BindingErrorHandling.ThrowException"/> (<see cref="TemplateErrorCode.PlaceholderNotReplaced"/> with
        /// the tag and its context, the cause as inner exception).
        /// </exception>
        public TextRenderResult Process()
        {
            var parsed = Parse();
            if (m_settings.BindingErrorHandling != BindingErrorHandling.HighlightErrorsInDocument)
            {
                ThrowIfSyntaxErrors();
            }
            var scriptCompiler = new ScriptCompiler(m_models, m_settings.ProcessSettings);
            return TextRenderer.Render(parsed, m_blockingErrors, m_settings, m_models, scriptCompiler);
        }

        private ParsedTextTemplate Parse()
        {
            if (m_parsed == null)
            {
                var processSettings = m_settings.ProcessSettings;
                var tree = TemplateSyntaxParser.Parse(m_template, PartName, processSettings, processSettings.TagContentPreprocessors);
                m_syntaxErrors = tree.Errors
                    .Concat(TextTemplateAnalyzer.FindWordOnlyConstructs(tree, m_template, PartName, processSettings))
                    .OrderBy(x => x.Index)
                    .ToList();
                m_blockingErrors = m_syntaxErrors.Where(x => x.Severity == TemplateSyntaxErrorSeverity.Error).ToList();
                m_parsed = new ParsedTextTemplate(m_template, tree, m_settings.IgnoreLineBreaksAroundTags);
            }
            return m_parsed;
        }

        private void ThrowIfSyntaxErrors()
        {
            if (m_blockingErrors.Count > 0)
            {
                throw TemplateException.Create(m_settings.ProcessSettings, TemplateErrorCode.TemplateSyntaxErrors, m_blockingErrors);
            }
        }
    }
}
