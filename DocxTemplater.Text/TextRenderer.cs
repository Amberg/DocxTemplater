using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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
        private readonly ITextTemplateContext m_context;
        private readonly TextValueFormatter m_valueFormatter;
        private readonly IReadOnlyList<(int Start, int End)> m_skip;
        private readonly StringBuilder m_output = new();
        private readonly List<string> m_errors = [];
        // the case selected by a switch node, set when the switch is rendered and read when its cases are rendered
        private readonly Dictionary<TemplateSyntaxNode, TemplateSyntaxNode> m_matchedCase = [];

        public TextRenderer(string text, TemplateSyntaxTree tree, ITextTemplateContext context, TextValueFormatter valueFormatter)
        {
            m_text = text;
            m_tree = tree;
            m_helper = new SyntaxTreeHelper(tree);
            m_context = context;
            m_valueFormatter = valueFormatter;
            m_skip = context.Settings.IgnoreLineBreaksAroundTags
                ? SyntaxTreeHelper.LineBreakRangesToSkip(text, tree.Matches)
                : Array.Empty<(int, int)>();
        }

        /// <summary>The errors collected with <see cref="BindingErrorHandling.HighlightErrorsInDocument"/>, in order of occurrence.</summary>
        public IReadOnlyList<string> Errors => m_errors;

        private ProcessSettings Settings => m_context.Settings.ProcessSettings;

        private BindingErrorHandling ErrorHandling => Settings.BindingErrorHandling;

        private bool Html => m_context.Settings.Escaping == TextEscaping.Html;

        public string Render()
        {
            RenderContent(m_tree.Blocks, 0, m_text.Length);
            return m_output.ToString();
        }

        /// <summary>
        /// The template with the tags of the <paramref name="errors"/> highlighted, not rendered - the counterpart
        /// of the Word renderer, which leaves a template with syntax errors as it is and marks the offending tags.
        /// </summary>
        public string HighlightSyntaxErrors(IReadOnlyList<TemplateSyntaxError> errors)
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
            AddError(TemplateErrorCode.TemplateSyntaxErrors, errors);
            return m_output.ToString();
        }

        /// <summary>The list of <see cref="Errors"/> as it is put in front of the output, like the red paragraph of the Word renderer.</summary>
        public string FormatErrorList()
        {
            var errors = m_errors.Distinct().ToList();
            if (errors.Count == 0)
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
                    ? new ValueWithMetadata(m_context.ScriptCompiler.CompileExpression(match.Variable)())
                    : m_context.ModelLookup.GetValueWithMetadata(match.Variable);
                AppendValue(m_valueFormatter.Format(m_context, match, value));
            }
            catch (Exception e) when (e is TemplateException or FormatException)
            {
                switch (ErrorHandling)
                {
                    case BindingErrorHandling.SkipBindingAndRemoveContent:
                        break;
                    case BindingErrorHandling.HighlightErrorsInDocument:
                        m_output.Append(Highlight(match.Match.Value));
                        m_errors.Add(e.Message);
                        break;
                    default:
                        int tagEnd = match.Index + match.Length;
                        throw TemplateException.Create(Settings, e, TemplateErrorCode.PlaceholderNotReplaced, match.Match.Value,
                            m_text[Math.Max(0, match.Index - ContextLength)..match.Index],
                            m_text[tagEnd..Math.Min(m_text.Length, tagEnd + ContextLength)]);
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
                model = m_context.ModelLookup.GetValue(collectionName);
            }
            catch (TemplateException e) when (ErrorHandling != BindingErrorHandling.ThrowException)
            {
                if (ErrorHandling == BindingErrorHandling.HighlightErrorsInDocument)
                {
                    m_errors.Add(e.Message);
                }
            }

            if (model is IEnumerable enumerable)
            {
                var body = node.Children[0];
                var separator = node.Children.Count > 1 ? node.Children[1] : null;
                var items = enumerable.Cast<object>().ToList();
                for (int i = 0; i < items.Count; i++)
                {
                    using var scope = m_context.ModelLookup.OpenScope();
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
            else if (model != null)
            {
                if (ErrorHandling == BindingErrorHandling.ThrowException)
                {
                    throw TemplateException.Create(Settings, TemplateErrorCode.NotEnumerable, collectionName, model.GetType().FullName);
                }
                AddError(TemplateErrorCode.NotEnumerable, collectionName, model.GetType().FullName);
            }
        }

        private void RenderCondition(TemplateSyntaxNode node)
        {
            var condition = node.Start.Condition;
            bool result = false;
            try
            {
                result = m_context.ScriptCompiler.CompileScript(condition)();
            }
            catch (TemplateException e) when (ErrorHandling != BindingErrorHandling.ThrowException)
            {
                AddError(TemplateErrorCode.ErrorInCondition, e, condition);
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
                    isMatch = m_context.ScriptCompiler.CompileScript($"{switchExpression} == {caseExpression}")();
                }
                catch (TemplateException e) when (ErrorHandling != BindingErrorHandling.ThrowException)
                {
                    AddError(TemplateErrorCode.ErrorInSwitchCase, e, switchExpression, caseExpression);
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
                count = RangeLoopCount.Resolve(m_context.ModelLookup, countVariable, Settings);
            }
            catch (TemplateException e) when (ErrorHandling != BindingErrorHandling.ThrowException)
            {
                if (ErrorHandling == BindingErrorHandling.HighlightErrorsInDocument)
                {
                    m_errors.Add(e.Message);
                }
            }

            for (int i = 0; i < count; i++)
            {
                using var scope = m_context.ModelLookup.OpenScope();
                scope.AddVariable(indexName, i);
                RenderSection(node.Children[0]);
            }
        }

        private void AddError(TemplateErrorCode code, params object[] arguments)
        {
            var messages = Settings.ErrorMessages ?? TemplateErrorMessages.Default;
            m_errors.Add(messages.Format(code, Settings.UiCulture, arguments));
        }

        /// <summary>A tag that could not be rendered: marked red in HTML, left as it is in plain text.</summary>
        private string Highlight(string tag)
        {
            return Html ? "<span style=\"" + HighlightStyle + "\">" + WebUtility.HtmlEncode(tag) + "</span>" : tag;
        }

        private static string EncodeWithLineBreaks(string text)
        {
            return WebUtility.HtmlEncode(text).Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "<br/>");
        }
    }
}
