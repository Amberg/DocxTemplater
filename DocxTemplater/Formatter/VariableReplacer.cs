using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocxTemplater.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater.Formatter
{
    internal class VariableReplacer : IVariableReplacer
    {
        private readonly IModelLookup m_models;
        private readonly List<IFormatter> m_formatters;
        private readonly List<string> m_errors;

        public VariableReplacer(IModelLookup models, ProcessSettings processSettings)
        {
            m_models = models;
            m_errors = new List<string>();
            ProcessSettings = processSettings;
            m_formatters = new List<IFormatter>();
            m_formatters.Add(new FormatPatternFormatter());
            m_formatters.Add(new HtmlFormatter());
            m_formatters.Add(new CaseFormatter());
        }

        public ProcessSettings ProcessSettings
        {
            get;
        }

        public void WriteErrorMessages(OpenXmlCompositeElement rootElement)
        {
            if (ProcessSettings.BindingErrorHandling != BindingErrorHandling.HighlightErrorsInDocument)
            {
                return;
            }

            if (rootElement is Document document)
            {
                rootElement = document.Body;
            }

            if (rootElement == null)
            {
                return;
            }

            if (m_errors.Count > 0)
            {
                var firstParagraph = rootElement.GetFirstChild<Paragraph>();
                var paragraph = new Paragraph();
                var text = new Text(string.Join("\n", m_errors.Distinct())) { Space = SpaceProcessingModeValues.Preserve };
                paragraph.AppendChild(new Run(new RunProperties()
                {
                    Color = new Color() { Val = "FF0000" },
                    Bold = new Bold()
                }, text));
                SplitNewLinesInText(text);
                if (firstParagraph != null)
                {
                    firstParagraph.InsertBeforeSelf(paragraph);
                }
                else
                {
                    rootElement.AddChild(paragraph);
                }
            }
        }

        public void AddError(string errorMessage)
        {
            m_errors.Add(errorMessage);
        }

        public void AddError(TemplateErrorCode errorCode, params object[] arguments)
        {
            var messages = ProcessSettings.ErrorMessages ?? TemplateErrorMessages.Default;
            m_errors.Add(messages.Format(errorCode, ProcessSettings.UiCulture, arguments));
        }

        public void RegisterFormatter(IFormatter formatter)
        {
            m_formatters.Add(formatter);
        }

        /// <summary>
        /// the formatter string is the leading formatter prefix, e.g. "FORMAT" followed by the formatter arguments ae image(100,200)
        /// </summary>
        private void ApplyFormatter(ITemplateProcessingContext templateContext, PatternMatch patternMatch, ValueWithMetadata valueWithMetadata, Text target)
        {
            var value = valueWithMetadata.Value;
            if (value == null)
            {
                target.Text = string.Empty;
                return;
            }

            var formatterText = GetFormatterText(patternMatch, valueWithMetadata, out string[] formatterArguments);
            ApplyFormatterInternal(templateContext, patternMatch, value, target, formatterText, formatterArguments);
        }

        internal void ApplyFormatter(PatternMatch patternMatch, object value, Text target, ITemplateProcessingContext templateContext)
        {
            if (value == null)
            {
                target.Text = string.Empty;
                return;
            }
            var formatterText = GetFormatterText(patternMatch, new ValueWithMetadata(value, new ValueMetadata()), out string[] formatterArguments);
            ApplyFormatterInternal(templateContext, patternMatch, value, target, formatterText, formatterArguments);
        }

        private void ApplyFormatterInternal(ITemplateProcessingContext templateContext, PatternMatch patternMatch, object value, Text target, string formatterText, string[] formatterArguments)
        {
            if (!string.IsNullOrWhiteSpace(formatterText))
            {
                foreach (var formatter in m_formatters)
                {
                    if (formatter.CanHandle(value.GetType(), formatterText))
                    {
                        var context = new FormatterContext(patternMatch.Variable, formatterText, formatterArguments,
                            value, ProcessSettings.Culture);
                        formatter.ApplyFormat(templateContext, context, target);
                        return;
                    }
                }
            }

            if (value is IFormattable formattable)
            {
                target.Text = formattable.ToString(null, ProcessSettings.Culture);
                return;
            }

            target.Text = value.ToString() ?? string.Empty;

        }

        public void ReplaceVariables(IReadOnlyCollection<OpenXmlElement> content,
            ITemplateProcessingContext templateContext)
        {
            foreach (var element in content)
            {
                ReplaceVariables(element, templateContext);
            }
        }

        public void ReplaceVariables(OpenXmlElement cloned, ITemplateProcessingContext templateContext)
        {
            // Fill content controls (structured document tags) whose tag is a placeholder, before the
            // text-placeholder pass. Runs from the same choke point so it applies to the root element and,
            // because loop bodies are rendered by cloning their content and calling this method again, to
            // every loop iteration as well.
            ReplaceContentControlValues(cloned, templateContext);

            var variables = cloned.GetElementsWithMarker(PatternType.Variable)
                .Concat(cloned.GetElementsWithMarker(PatternType.Expression))
                .OfType<Text>().ToList();
            foreach (var text in variables)
            {
                var variableMatch = PatternMatcher.FindSyntaxPatterns(text.Text, ProcessSettings).FirstOrDefault() ??
                                    throw OpenXmlTemplateException.Create(ProcessSettings, TemplateErrorCode.InvalidVariableSyntax, text.Text);
                try
                {
                    if (variableMatch.Type == PatternType.Expression)
                    {
                        var expressionResult = templateContext.ScriptCompiler.CompileExpression(variableMatch.Variable)();
                        ApplyFormatter(variableMatch, expressionResult, text, templateContext);
                    }
                    else
                    {
                        var valueWithMetadata = m_models.GetValueWithMetadata(variableMatch.Variable);
                        ApplyFormatter(templateContext, variableMatch, valueWithMetadata, text);
                    }
                    VariableReplacer.SplitNewLinesInText(text);
                }
                catch (Exception e) when (e is OpenXmlTemplateException or FormatException)
                {
                    if (ProcessSettings.BindingErrorHandling == BindingErrorHandling.SkipBindingAndRemoveContent)
                    {
                        text.RemoveWithEmptyParent();
                    }
                    else if (ProcessSettings.BindingErrorHandling == BindingErrorHandling.HighlightErrorsInDocument)
                    {
                        MarkTextAsError(text);
                        AddError(e.Message);
                    }
                    else
                    {
                        throw OpenXmlTemplateException.Create(ProcessSettings, e, TemplateErrorCode.PlaceholderNotReplaced,
                            text.InnerText, text.ElementBeforeInDocument<Text>()?.InnerText, text.ElementAfterInDocument<Text>()?.InnerText);
                    }
                }
            }
        }

        /// <summary>
        /// Fills Word content controls (structured document tags) whose <c>w:tag</c> is a DocxTemplater
        /// placeholder (e.g. a tag of <c>{{ds.Name}}</c> or <c>{{ds.Price}:f(c)}</c>). The resolved model
        /// value replaces the control's content using the same model lookup and formatters as text
        /// placeholders, and the "showing placeholder" flag is cleared so Word treats the value as real text.
        /// </summary>
        /// <remarks>
        /// Unlike a text placeholder, a content control is a named, persistent region: it is never removed,
        /// and its tag is left untouched, so the control keeps its identity and can be filled again in a
        /// later processing pass (e.g. stamping a delivery date after the document was generated and reviewed).
        /// Tags that are not a placeholder - or whose placeholder is a block directive such as
        /// <c>{{#items}}</c> - are ignored, so existing templates are unaffected.
        /// </remarks>
        private void ReplaceContentControlValues(OpenXmlElement root, ITemplateProcessingContext templateContext)
        {
            // Opt-in: content-control tag binding inspects the w:tag of pre-existing controls, which could
            // change behavior of templates not authored for DocxTemplater. It is off unless explicitly enabled.
            if (!ProcessSettings.EnableContentControlTagBinding)
            {
                return;
            }

            var contentControls = root.Descendants<SdtElement>().ToList();
            if (root is SdtElement rootControl)
            {
                contentControls.Insert(0, rootControl);
            }

            foreach (var contentControl in contentControls)
            {
                var tag = contentControl.SdtProperties?.GetFirstChild<Tag>()?.Val?.Value;
                if (string.IsNullOrEmpty(tag) || !tag.Contains("{{"))
                {
                    continue;
                }

                PatternMatch match;
                try
                {
                    match = PatternMatcher.FindSyntaxPatterns(tag, ProcessSettings).FirstOrDefault();
                }
                catch (OpenXmlTemplateException)
                {
                    // The tag contains braces but is a malformed directive (e.g. "{{@}}" or "{{/switch}}"),
                    // not a value binding. Treat it like any non-placeholder tag and leave the control
                    // untouched, rather than letting the parse error abort the whole render.
                    continue;
                }

                if (match == null || (match.Type != PatternType.Variable && match.Type != PatternType.Expression))
                {
                    continue;
                }

                FillContentControl(contentControl, tag, match, templateContext);
            }
        }

        private void FillContentControl(SdtElement contentControl, string tag, PatternMatch match, ITemplateProcessingContext templateContext)
        {
            // Resolve the value BEFORE touching the control's content. If the tag cannot be bound, the
            // control is left exactly as it was - see ApplyContentControlErrorMode. This is what makes
            // multi-pass filling safe: a later pass that does not bind this tag must not clear a value a
            // previous pass wrote.
            object value;
            try
            {
                value = match.Type == PatternType.Expression
                    ? templateContext.ScriptCompiler.CompileExpression(match.Variable)()
                    : m_models.GetValueWithMetadata(match.Variable);
            }
            catch (Exception e) when (e is OpenXmlTemplateException or FormatException)
            {
                ApplyContentControlErrorMode(contentControl, tag, e, preparedTarget: null);
                return;
            }

            // A resolved-but-null value (a bound member that is simply null, as opposed to an absent one
            // that throws above) is treated like an unbound tag: the control is left untouched. This keeps
            // multi-pass filling non-destructive when passes share a model type whose other members are null.
            var resolvedValue = match.Type == PatternType.Expression ? value : ((ValueWithMetadata)value).Value;
            if (resolvedValue == null)
            {
                return;
            }

            var target = PrepareContentControlTarget(contentControl);
            if (target == null)
            {
                // The value resolved but there is nowhere to put it (e.g. an empty cell/row content control).
                // Surface this per the error mode instead of silently dropping the resolved value.
                ApplyContentControlErrorMode(contentControl, tag,
                    OpenXmlTemplateException.Create(ProcessSettings, TemplateErrorCode.ContentControlHasNoText),
                    preparedTarget: null);
                return;
            }

            try
            {
                if (match.Type == PatternType.Expression)
                {
                    ApplyFormatter(match, value, target, templateContext);
                }
                else
                {
                    ApplyFormatter(templateContext, match, (ValueWithMetadata)value, target);
                }
                SplitNewLinesInText(target);
            }
            catch (Exception e) when (e is OpenXmlTemplateException or FormatException)
            {
                ApplyContentControlErrorMode(contentControl, tag, e, target);
                return;
            }

            RemoveShowingPlaceholder(contentControl);
        }

        /// <summary>
        /// Applies the configured <see cref="BindingErrorHandling"/> to a content control whose tag could
        /// not be resolved or formatted, without ever removing the control itself.
        /// <paramref name="preparedTarget"/> is <c>null</c> when the failure happened before the content was
        /// touched (value resolution): in that case the control is left completely unchanged under
        /// <see cref="BindingErrorHandling.SkipBindingAndRemoveContent"/>.
        /// </summary>
        private void ApplyContentControlErrorMode(SdtElement contentControl, string tag, Exception e, Text preparedTarget)
        {
            switch (ProcessSettings.BindingErrorHandling)
            {
                case BindingErrorHandling.SkipBindingAndRemoveContent:
                    if (preparedTarget is { } clearedTarget)
                    {
                        clearedTarget.Text = string.Empty;
                    }
                    break;
                case BindingErrorHandling.HighlightErrorsInDocument:
                    // For a resolution failure (no prepared target) highlight the control's existing content
                    // in place WITHOUT clearing it, so the marker is visible and the original text is kept,
                    // mirroring how a text placeholder is highlighted. For a formatting failure the target is
                    // already prepared.
                    var target = preparedTarget ?? FirstOwnText(contentControl);
                    if (target != null)
                    {
                        MarkTextAsError(target);
                    }
                    AddError(e.Message);
                    break;
                default:
                    throw OpenXmlTemplateException.Create(ProcessSettings, e, TemplateErrorCode.ContentControlNotReplaced, tag, e);
            }
        }

        /// <summary>
        /// Reduces the content control's own content to a single, empty, unmarked <see cref="Text"/> that
        /// the formatter pipeline can write into. The first existing text run is reused (so its formatting
        /// is kept); any further text runs - and the now-empty runs/paragraphs around them - are dropped, so
        /// no blank lines are left behind. Text belonging to a nested content control is never touched. An
        /// empty inline/block control gets a minimal run so there is something to fill. Returns <c>null</c>
        /// when the control has no place to put text (e.g. an empty cell/row control), leaving it untouched.
        /// </summary>
        private static Text PrepareContentControlTarget(SdtElement contentControl)
        {
            var content = contentControl.ChildElements
                .FirstOrDefault(c => c is SdtContentBlock or SdtContentRun or SdtContentCell or SdtContentRow);
            if (content == null)
            {
                return null;
            }

            var texts = OwnTexts(contentControl, content);
            if (texts.Count > 0)
            {
                var target = texts[0];
                for (var i = 1; i < texts.Count; i++)
                {
                    texts[i].RemoveWithEmptyParent();
                }
                // Drop a marker left by the isolate pass if the control's placeholder text happened to
                // contain '{{...}}', so the text-placeholder pass does not process (and possibly remove) it.
                target.RemoveMarker();
                target.Text = string.Empty;
                target.Space = SpaceProcessingModeValues.Preserve;
                return target;
            }

            var newText = new Text { Space = SpaceProcessingModeValues.Preserve };
            if (content is SdtContentRun)
            {
                content.AppendChild(new Run(newText));
                return newText;
            }
            if (content is SdtContentBlock)
            {
                content.AppendChild(new Paragraph(new Run(newText)));
                return newText;
            }
            return null;
        }

        /// <summary>
        /// The <see cref="Text"/> nodes that belong directly to this content control - i.e. not to a
        /// content control nested inside it. Scoping to the control's own text keeps an outer control from
        /// consuming or corrupting the content of an inner one.
        /// </summary>
        private static List<Text> OwnTexts(SdtElement contentControl, OpenXmlElement content)
        {
            return content.Descendants<Text>()
                .Where(t => t.Ancestors<SdtElement>().FirstOrDefault() == contentControl)
                .ToList();
        }

        private static Text FirstOwnText(SdtElement contentControl)
        {
            var content = contentControl.ChildElements
                .FirstOrDefault(c => c is SdtContentBlock or SdtContentRun or SdtContentCell or SdtContentRow);
            return content == null ? null : OwnTexts(contentControl, content).FirstOrDefault();
        }

        private static void RemoveShowingPlaceholder(SdtElement contentControl)
        {
            contentControl.SdtProperties?.GetFirstChild<ShowingPlaceholder>()?.Remove();
        }

        /// <summary>
        /// Use the formatter from the template, if not available use the default formatter from the metadata
        /// set through <see cref="ModelPropertyAttribute"/>
        /// </summary>
        private string GetFormatterText(PatternMatch patternMatch, ValueWithMetadata valueWithMetadata,
            out string[] formatterArguments)
        {
            formatterArguments = patternMatch.Arguments;
            var formatterText = patternMatch.Formatter;
            if (string.IsNullOrWhiteSpace(formatterText))
            {
                if (!string.IsNullOrWhiteSpace(valueWithMetadata.Metadata.DefaultFormatter))
                {
                    // try to parse default formatter from metadata
                    var found = PatternMatcher
                        .FindSyntaxPatterns("{{x}:" + valueWithMetadata.Metadata.DefaultFormatter + "}", ProcessSettings)
                        .FirstOrDefault();
                    if (found != null && !string.IsNullOrWhiteSpace(found.Formatter))
                    {
                        formatterText = found.Formatter;
                        formatterArguments = found.Arguments;
                    }
                }
            }

            return formatterText;
        }

        /// <summary>
        /// Insert Breaks for line breaks in the text
        /// </summary>
        internal static void SplitNewLinesInText(Text text)
        {
            if (text.Parent == null)
            {
                return;
            }

            text.Text = text.Text.Replace("\r\n", "\n").Replace("\r", "\n");
            if (text.Text.Contains('\n'))
            {
                // Emit exactly one break per '\n' (i.e. one break between consecutive segments).
                // A break is therefore only produced for a newline that is actually present in the
                // value: "a\nb" -> a<br>b, "a\n" -> a<br>, "a\n\nb" -> a<br><br>b.
                // The previous implementation added a break after every non-empty segment, which
                // produced a spurious trailing break for multi-line values without a trailing
                // newline (rendered as an empty line / extra blank table row).
                var parts = text.Text.Split('\n');
                OpenXmlElement lastElement = text;
                for (var i = 0; i < parts.Length; i++)
                {
                    if (i > 0)
                    {
                        lastElement = lastElement.InsertAfterSelf(new Break());
                    }

                    if (!string.IsNullOrWhiteSpace(parts[i]))
                    {
                        lastElement = lastElement.InsertAfterSelf(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
                    }
                }

                text.Remove();
            }
        }

        public static void MarkTextAsError(Text text)
        {
            var run = text.GetFirstAncestor<Run>();
            if (run != null)
            {
                // Get or create run properties
                var runProperties = run.GetFirstChild<RunProperties>();
                if (runProperties == null)
                {
                    runProperties = new RunProperties();
                    run.InsertAt(runProperties, 0);
                }

                // Set text color to black and background color to red
                runProperties.Color = new Color() { Val = "000000" }; // Black text
                runProperties.Bold = new Bold();

                // Add red background
                var shading = new Shading()
                {
                    Val = ShadingPatternValues.Clear,  // Background shading
                    Color = "auto",                    // Automatic text color
                    Fill = "FF0000"                     // Red background
                };
                runProperties.AddChild(shading);

                text.RemoveMarker();
            }
        }
    }
}
