using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocxTemplater.Blocks;

namespace DocxTemplater.Text
{
    /// <summary>
    /// Renders the syntax tree of a text template into a string. Walks the tree in document order: the literal text
    /// between tags is copied, values are resolved and encoded, blocks are expanded with the same semantics as the
    /// blocks of the Word renderer (scope variables, <c>_Idx</c>, separators, error handling).
    /// </summary>
    internal sealed class TextRenderer
    {
        private const int ContextLength = 30;
        private const string HighlightStyle = "background-color:#ff0000;color:#000000;font-weight:bold";
        private const string ErrorListStyle = "color:#ff0000;font-weight:bold";

        private readonly string m_text;
        private readonly TemplateSyntaxTree m_tree;
        private readonly SyntaxTreeHelper m_helper;
        private readonly TextTemplateSettings m_settings;
        private readonly IModelLookup m_models;
        private readonly IScriptCompiler m_scriptCompiler;
        private readonly IReadOnlyList<(int Start, int End)> m_skip;
        private readonly StringBuilder m_output;
        private readonly List<TextRenderError> m_errors = [];
        // the case selected by a switch node, set when the switch is rendered and read when its cases are rendered
        private readonly Dictionary<TemplateSyntaxNode, TemplateSyntaxNode> m_matchedCase = [];

        private TextRenderer(ParsedTextTemplate template, TextTemplateSettings settings, IModelLookup models, IScriptCompiler scriptCompiler)
        {
            m_text = template.Text;
            m_tree = template.Tree;
            m_helper = template.Helper;
            m_skip = template.SkippedRanges;
            m_settings = settings;
            m_models = models;
            m_scriptCompiler = scriptCompiler;
            // the output is usually about as long as the template; one buffer avoids growing it chunk by chunk
            m_output = new StringBuilder(m_text.Length + (m_text.Length / 4));
        }

        /// <summary>
        /// Renders <paramref name="template"/> - or, if it has <paramref name="syntaxErrors"/>, returns it with the
        /// offending tags highlighted. With <see cref="BindingErrorHandling.HighlightErrorsInDocument"/> the errors are
        /// listed at the top of the text; in all modes that do not throw they are returned in the result.
        /// </summary>
        public static TextRenderResult Render(ParsedTextTemplate template, IReadOnlyList<TemplateSyntaxError> syntaxErrors,
            TextTemplateSettings settings, IModelLookup models, IScriptCompiler scriptCompiler)
        {
            var renderer = new TextRenderer(template, settings, models, scriptCompiler);
            var output = syntaxErrors.Count > 0 ? renderer.HighlightSyntaxErrors(syntaxErrors) : renderer.Render();
            return new TextRenderResult(renderer.FormatErrorList() + output, renderer.m_errors);
        }

        private ProcessSettings Settings => m_settings.ProcessSettings;

        private BindingErrorHandling ErrorHandling => Settings.BindingErrorHandling;

        private bool Html => m_settings.Escaping == TextEscaping.Html;

        private string Render()
        {
            RenderContent(m_tree.Blocks, 0, m_text.Length);
            return m_output.ToString();
        }

        /// <summary>
        /// The template with the tags of the <paramref name="errors"/> highlighted, not rendered - the counterpart
        /// of the Word renderer, which leaves a template with syntax errors as it is and marks the offending tags.
        /// </summary>
        private string HighlightSyntaxErrors(IReadOnlyList<TemplateSyntaxError> errors)
        {
            int pos = 0;
            foreach (var error in errors.Where(x => x.Length > 0).OrderBy(x => x.Index))
            {
                // overlapping ranges (e.g. two errors on the same tag) are highlighted once
                int first = Math.Max(error.Index, pos);
                int end = Math.Min(error.Index + error.Length, m_text.Length);
                if (end <= first)
                {
                    continue;
                }
                m_output.Append(m_text, pos, first - pos);
                m_output.Append(Highlight(m_text[first..end]));
                pos = end;
            }
            m_output.Append(m_text, pos, m_text.Length - pos);
            m_errors.Add(new TextRenderError(string.Empty, 0, 0, TemplateException.Create(Settings, TemplateErrorCode.TemplateSyntaxErrors, errors)));
            return m_output.ToString();
        }

        /// <summary>
        /// The list of errors as it is put in front of the output with <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>,
        /// like the red paragraph of the Word renderer; empty in the other modes.
        /// </summary>
        private string FormatErrorList()
        {
            var errors = m_errors.Select(x => x.Exception.Message).Distinct().ToList();
            if (errors.Count == 0 || ErrorHandling != BindingErrorHandling.HighlightErrorsInDocument)
            {
                return string.Empty;
            }
            if (Html)
            {
                return "<div style=\"" + ErrorListStyle + "\">" + string.Join("<br/>", errors.Select(EncodeWithLineBreaks)) + "</div>";
            }
            return string.Join("\n", errors) + "\n";
        }

        /// <summary>Renders <c>[from, to)</c> of the template: the literal text and values between the given child nodes, and the nodes.</summary>
        private void RenderContent(IReadOnlyList<TemplateSyntaxNode> children, int from, int to)
        {
            int pos = from;
            foreach (var child in children)
            {
                RenderLiteralAndValues(pos, child.Start.Index);
                RenderNode(child);
                pos = SyntaxTreeHelper.NodeEnd(child);
            }
            RenderLiteralAndValues(pos, to);
        }

        private void RenderLiteralAndValues(int from, int to)
        {
            int pos = from;
            foreach (var value in m_helper.ValuesIn(from, to))
            {
                AppendLiteral(pos, value.Index);
                RenderValue(value);
                pos = value.Index + value.Length;
            }
            AppendLiteral(pos, to);
        }

        /// <summary>Copies template text, leaving out the ranges removed by <see cref="TextTemplateSettings.IgnoreLineBreaksAroundTags"/>.</summary>
        private void AppendLiteral(int from, int to)
        {
            int pos = from;
            foreach (var (start, end) in m_skip)
            {
                if (end <= pos)
                {
                    continue;
                }
                if (start >= to)
                {
                    break;
                }
                m_output.Append(m_text, pos, Math.Max(start, pos) - pos);
                pos = Math.Min(end, to);
            }
            m_output.Append(m_text, pos, to - pos);
        }

        private void RenderSection(TemplateSyntaxNode section)
        {
            RenderContent(section.Children, SyntaxTreeHelper.SectionStart(section), SyntaxTreeHelper.SectionEnd(section));
        }

        private void RenderNode(TemplateSyntaxNode node)
        {
            switch (node.Type)
            {
                case PatternType.InlineKeyWord:
                    // only {{:break}} gets here - the Word-only keywords are rejected by the validation
                    m_output.Append(Html ? "<br/>" : "\n");
                    break;
                case PatternType.IgnoreBlock:
                    AppendLiteral(SyntaxTreeHelper.SectionStart(node), SyntaxTreeHelper.SectionEnd(node));
                    break;
                case PatternType.Condition:
                    RenderCondition(node);
                    break;
                case PatternType.CollectionStart:
                    RenderLoop(node);
                    break;
                case PatternType.RangeStart:
                    RenderRangeLoop(node);
                    break;
                case PatternType.Switch:
                    RenderSwitch(node);
                    break;
                case PatternType.Case:
                case PatternType.Default:
                    // the case's parent is the content node of the switch, whose parent is the switch node
                    if (m_matchedCase.TryGetValue(node.Parent.Parent, out var matched) && matched == node)
                    {
                        RenderSection(node.Children[0]);
                    }
                    break;
            }
        }

        private void RenderValue(PatternMatch match)
        {
            try
            {
                var value = match.Type == PatternType.Expression
                    ? new ValueWithMetadata(m_scriptCompiler.CompileExpression(match.Variable)())
                    : m_models.GetValueWithMetadata(match.Variable);
                AppendValue(m_settings.ValueFormatter.Format(m_settings, match, value));
            }
            catch (Exception e)
            {
                Fail(match, AsTemplateException(e));
                if (ErrorHandling == BindingErrorHandling.HighlightErrorsInDocument)
                {
                    m_output.Append(Highlight(match.Match.Value));
                }
            }
        }

        private void AppendValue(TextValue value)
        {
            m_output.Append(value.IsRaw || !Html ? value.Text : EncodeWithLineBreaks(value.Text));
        }

        private void RenderLoop(TemplateSyntaxNode node)
        {
            var collectionName = node.Start.Variable;
            object model = null;
            try
            {
                model = m_models.GetValue(collectionName);
                if (model is not (null or IEnumerable))
                {
                    throw TemplateException.Create(Settings, TemplateErrorCode.NotEnumerable, collectionName, model.GetType().FullName);
                }
            }
            catch (Exception e)
            {
                Fail(node.Start, AsTemplateException(e));
            }

            if (model is IEnumerable enumerable)
            {
                var body = node.Children[0];
                var separator = node.Children.Count > 1 ? node.Children[1] : null;
                var items = enumerable.Cast<object>().ToList();
                for (int i = 0; i < items.Count; i++)
                {
                    using var scope = m_models.OpenScope();
                    scope.AddVariable(collectionName, items[i]);
                    scope.AddVariable($"{collectionName}.{LoopBlock.LoopIndexVariable}", i + 1);
                    scope.AddVariable($"{collectionName}.{LoopBlock.LoopLengthVariable}", items.Count);
                    RenderSection(body);
                    if (separator != null && i < items.Count - 1)
                    {
                        RenderSection(separator);
                    }
                }
            }
        }

        private void RenderCondition(TemplateSyntaxNode node)
        {
            var condition = node.Start.Condition;
            bool result = false;
            try
            {
                result = m_scriptCompiler.CompileScript(condition)();
            }
            catch (Exception e)
            {
                Fail(node.Start, TemplateException.Create(Settings, e, TemplateErrorCode.ErrorInCondition, e, condition));
            }

            if (result)
            {
                RenderSection(node.Children[0]);
            }
            else if (node.Children.Count > 1)
            {
                RenderSection(node.Children[1]);
            }
        }

        private void RenderSwitch(TemplateSyntaxNode node)
        {
            var switchExpression = SyntaxTreeHelper.KeywordArgument(node.Start);
            var content = node.Children[0];
            TemplateSyntaxNode matched = null;
            TemplateSyntaxNode defaultCase = null;
            foreach (var caseNode in content.Children.Where(x => x.Type is PatternType.Case or PatternType.Default))
            {
                if (caseNode.Type == PatternType.Default)
                {
                    defaultCase = caseNode;
                    continue;
                }
                if (matched != null)
                {
                    continue;
                }
                var caseExpression = SyntaxTreeHelper.KeywordArgument(caseNode.Start);
                bool isMatch = false;
                try
                {
                    isMatch = m_scriptCompiler.CompileScript($"{switchExpression} == {caseExpression}")();
                }
                catch (Exception e)
                {
                    Fail(caseNode.Start, TemplateException.Create(Settings, e, TemplateErrorCode.ErrorInSwitchCase, e, switchExpression, caseExpression));
                }
                if (isMatch)
                {
                    matched = caseNode;
                }
            }
            m_matchedCase[node] = matched ?? defaultCase;
            RenderSection(content);
        }

        private void RenderRangeLoop(TemplateSyntaxNode node)
        {
            var (indexName, countVariable) = RangeLoopCount.Parse(node.Start.Variable);
            int count = 0;
            try
            {
                count = RangeLoopCount.Resolve(m_models, countVariable, Settings);
            }
            catch (Exception e)
            {
                Fail(node.Start, AsTemplateException(e));
            }

            for (int i = 0; i < count; i++)
            {
                using var scope = m_models.OpenScope();
                scope.AddVariable(indexName, i);
                RenderSection(node.Children[0]);
            }
        }

        /// <summary>
        /// A tag that could not be rendered: with <see cref="BindingErrorHandling.ThrowException"/> thrown as
        /// <see cref="TemplateErrorCode.PlaceholderNotReplaced"/> with the tag and its context, otherwise collected.
        /// </summary>
        private void Fail(PatternMatch match, TemplateException error)
        {
            if (ErrorHandling == BindingErrorHandling.ThrowException)
            {
                int tagEnd = match.Index + match.Length;
                throw TemplateException.Create(Settings, error, TemplateErrorCode.PlaceholderNotReplaced, match.Match.Value,
                    m_text[Math.Max(0, match.Index - ContextLength)..match.Index],
                    m_text[tagEnd..Math.Min(m_text.Length, tagEnd + ContextLength)]);
            }
            m_errors.Add(new TextRenderError(match.Match.Value, match.Index, match.Length, error));
        }

        // an exception of the model, an expression (division by zero, index) or a formatter is an error of the tag
        private static TemplateException AsTemplateException(Exception e)
        {
            return e as TemplateException ?? new TemplateException(e.Message, e);
        }

        /// <summary>A tag that could not be rendered: marked red in HTML, left as it is in plain text.</summary>
        private string Highlight(string tag)
        {
            return Html ? "<span style=\"" + HighlightStyle + "\">" + HtmlEncoding.Encode(tag) + "</span>" : tag;
        }

        private string EncodeWithLineBreaks(string text)
        {
            return HtmlEncoding.Encode(text).Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "<br/>");
        }
    }
}
