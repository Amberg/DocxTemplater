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
    /// The template is parsed once; <see cref="Process"/> can be called repeatedly, e.g. to render a mail for
    /// several recipients with a different model each time. Constructs that insert Word content
    /// (<c>:img()</c>, <c>:template()</c>, <c>:md()</c>, <c>:dyntable()</c>, <c>{{:PageBreak}}</c>,
    /// <c>{{:SectionBreak}}</c>) are reported as syntax errors. An instance is not thread-safe.
    /// </remarks>
    public sealed class TextTemplate
    {
        /// <summary>The <see cref="TemplateSyntaxError.Part"/> of the errors of a text template.</summary>
        public const string PartName = "Text";

        private readonly ModelLookup m_modelLookup;
        private readonly TextTemplateContext m_context;
        private readonly TextValueFormatter m_valueFormatter = new();
        private TemplateSyntaxTree m_tree;
        private TemplateSchema m_schema;

        public TextTemplate(string template)
            : this(template, TextTemplateSettings.Default)
        {
        }

        public TextTemplate(string template, TextTemplateSettings settings)
        {
            ArgumentNullException.ThrowIfNull(template);
            ArgumentNullException.ThrowIfNull(settings);
            Template = template;
            Settings = settings;
            m_modelLookup = new ModelLookup(settings.ProcessSettings);
            m_context = new TextTemplateContext(settings, m_modelLookup, new ScriptCompiler(m_modelLookup, settings.ProcessSettings));
        }

        public TextTemplateSettings Settings { get; }

        /// <summary>The text of the template as passed to the constructor.</summary>
        public string Template { get; }

        /// <summary>The models bound with <see cref="BindModel"/>, by prefix.</summary>
        public IReadOnlyDictionary<string, object> Models => m_modelLookup.Models;

        /// <summary>
        /// The errors of the last <see cref="Process"/> with <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>,
        /// as they are listed at the top of the output; empty if the template rendered without errors.
        /// </summary>
        public IReadOnlyList<string> Errors { get; private set; } = Array.Empty<string>();

        // parsed on first use, so the error messages are in the language the settings have at that time
        private TemplateSyntaxTree GetTree()
        {
            return m_tree ??= TemplateSyntaxParser.Parse(Template, PartName, Settings.ProcessSettings);
        }

        /// <summary>
        /// Binds a model under a prefix: <c>{{prefix.Property}}</c>. The prefix of the first model can be omitted in the
        /// template, <c>{{Property}}</c>.
        /// </summary>
        public void BindModel(string prefix, object model)
        {
            m_modelLookup.Add(prefix, model);
        }

        public void RegisterFormatter(ITextFormatter formatter)
        {
            m_valueFormatter.Register(formatter);
        }

        /// <summary>
        /// Checks the template syntax without rendering it: that all blocks are closed, closing tags match their
        /// opening tags, else / separator / case tags are used inside the right block, tags are well-formed and no
        /// construct is used that exists only in Word templates. Binding errors (unknown variables, wrong types)
        /// are not detected - they require a model.
        /// </summary>
        /// <returns>
        /// All syntax errors found, ordered by position; an empty list if the syntax is valid. Entries with
        /// <see cref="TemplateSyntaxErrorSeverity.Error"/> make <see cref="Process"/> throw; warnings do not.
        /// </returns>
        public IReadOnlyList<TemplateSyntaxError> ValidateTemplateSyntax()
        {
            var tree = GetTree();
            return tree.Errors
                .Concat(TextTemplateAnalyzer.FindWordOnlyConstructs(tree, Template, PartName, Settings.ProcessSettings))
                .OrderBy(x => x.Index)
                .ToList();
        }

        /// <summary>
        /// Statically analyzes the template (without rendering it) and returns the structural schema of the
        /// variables, collections and nested objects the template references - the union of all branches, see
        /// <see cref="TemplateSchema"/>. Can be called before or after <see cref="Process"/>; the result is cached.
        /// </summary>
        /// <exception cref="TemplateException">The template has syntax errors.</exception>
        public TemplateSchema GetTemplateSchema()
        {
            if (m_schema == null)
            {
                ThrowIfSyntaxErrors(BlockingErrors());
                m_schema = TextTemplateAnalyzer.CollectSchema(GetTree(), Template.Length);
            }
            return m_schema;
        }

        /// <summary>
        /// Renders the template with the bound models.
        /// </summary>
        /// <exception cref="TemplateException">
        /// The template has syntax errors, or a placeholder could not be bound and
        /// <see cref="TextTemplateSettings.BindingErrorHandling"/> is <see cref="BindingErrorHandling.ThrowException"/>.
        /// With <see cref="BindingErrorHandling.HighlightErrorsInDocument"/> a template with syntax errors is not
        /// rendered: the offending tags are highlighted and the errors are listed at the top, like binding errors.
        /// </exception>
        public string Process()
        {
            var renderer = new TextRenderer(Template, GetTree(), m_context, m_valueFormatter);
            var syntaxErrors = BlockingErrors();
            string output;
            if (syntaxErrors.Count > 0)
            {
                if (Settings.BindingErrorHandling != BindingErrorHandling.HighlightErrorsInDocument)
                {
                    ThrowIfSyntaxErrors(syntaxErrors);
                }
                output = renderer.HighlightSyntaxErrors(syntaxErrors);
            }
            else
            {
                output = renderer.Render();
            }
            Errors = renderer.Errors;
            return renderer.FormatErrorList() + output;
        }

        private List<TemplateSyntaxError> BlockingErrors()
        {
            return ValidateTemplateSyntax().Where(x => x.Severity == TemplateSyntaxErrorSeverity.Error).ToList();
        }

        private void ThrowIfSyntaxErrors(List<TemplateSyntaxError> errors)
        {
            if (errors.Count > 0)
            {
                throw TemplateException.Create(Settings.ProcessSettings, TemplateErrorCode.TemplateSyntaxErrors, errors);
            }
        }
    }
}
