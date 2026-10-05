using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocxTemplater
{
    /// <summary>
    /// A block of the template syntax tree. The tree has the same shape as the <see cref="Blocks.ContentBlock"/> tree
    /// built from it: every opening tag yields a node with a child <see cref="PatternType.None"/> node holding the content;
    /// else / separator tags yield siblings of that content node.
    /// </summary>
    internal sealed class TemplateSyntaxNode
    {
        private readonly List<TemplateSyntaxNode> m_children = [];

        public TemplateSyntaxNode(PatternType type, PatternMatch start, TemplateSyntaxNode parent)
        {
            Type = type;
            Start = start;
            Parent = parent;
            parent?.m_children.Add(this);
        }

        public PatternType Type { get; }

        public PatternMatch Start { get; }

        public PatternMatch End { get; set; }

        public TemplateSyntaxNode Parent { get; }

        public IReadOnlyList<TemplateSyntaxNode> Children => m_children;
    }

    internal sealed class TemplateSyntaxTree
    {
        private readonly ProcessSettings m_settings;

        public TemplateSyntaxTree(IReadOnlyList<PatternMatch> matches, IReadOnlyList<TemplateSyntaxNode> blocks, IReadOnlyList<TemplateSyntaxError> errors, ProcessSettings settings)
        {
            Matches = matches;
            Blocks = blocks;
            Errors = errors;
            m_settings = settings;
        }

        /// <summary>All valid syntax patterns in document order.</summary>
        public IReadOnlyList<PatternMatch> Matches { get; }

        /// <summary>The top level blocks.</summary>
        public IReadOnlyList<TemplateSyntaxNode> Blocks { get; }

        public IReadOnlyList<TemplateSyntaxError> Errors { get; }

        public bool HasErrors => Errors.Any(x => x.Severity == TemplateSyntaxErrorSeverity.Error);

        public void ThrowIfErrors()
        {
            if (HasErrors)
            {
                var errors = Errors.Where(x => x.Severity == TemplateSyntaxErrorSeverity.Error).ToList();
                throw OpenXmlTemplateException.Create(m_settings, TemplateErrorCode.TemplateSyntaxErrors, errors);
            }
        }
    }

    /// <summary>
    /// Parses the text of a template part into a <see cref="TemplateSyntaxTree"/>. All syntax errors are collected
    /// instead of throwing on the first one, so the same parser serves <see cref="DocxTemplate.ValidateTemplateSyntax"/>
    /// and the rendering pipeline.
    /// </summary>
    internal static class TemplateSyntaxParser
    {
        private const int ContextLength = 30;
        private const int MaxUnterminatedTagLength = 40;

        // remains of a tag that the pattern matcher did not recognize, e.g. "{{Name}" or "{{first name}}"
        private static readonly Regex StrayBracesRegex = new(@"(?<open>\{\s*(?:\?\s*)?\{)|(?<close>\}\s*\})", RegexOptions.Compiled, TimeSpan.FromSeconds(1));

        private static readonly HashSet<string> InlineKeyWords = new(StringComparer.OrdinalIgnoreCase) { "Break", "PageBreak", "SectionBreak" };

        private sealed class ParseContext
        {
            private readonly string m_text;
            private readonly string m_part;
            private readonly ProcessSettings m_settings;
            private readonly List<(int Index, TemplateSyntaxError Error)> m_errors = [];

            public ParseContext(string text, string part, ProcessSettings settings)
            {
                m_text = text;
                m_part = part;
                m_settings = settings;
            }

            public void AddError(int index, int length, TemplateErrorCode code, params object[] args)
            {
                Add(TemplateSyntaxErrorSeverity.Error, index, length, code, args);
            }

            public void AddError(PatternMatch match, TemplateErrorCode code, params object[] args)
            {
                Add(TemplateSyntaxErrorSeverity.Error, match.Index, match.Length, code, args);
            }

            public void AddWarning(PatternMatch match, TemplateErrorCode code, params object[] args)
            {
                Add(TemplateSyntaxErrorSeverity.Warning, match.Index, match.Length, code, args);
            }

            public void AddWarning(int index, int length, TemplateErrorCode code, params object[] args)
            {
                Add(TemplateSyntaxErrorSeverity.Warning, index, length, code, args);
            }

            public IReadOnlyList<TemplateSyntaxError> SortedErrors => m_errors.OrderBy(x => x.Index).Select(x => x.Error).ToList();

            private void Add(TemplateSyntaxErrorSeverity severity, int index, int length, TemplateErrorCode code, object[] args)
            {
                var start = Math.Max(0, index - ContextLength);
                var end = Math.Min(m_text.Length, index + length + ContextLength);
                m_errors.Add((index, new TemplateSyntaxError(severity, m_part, m_text.Substring(index, length), m_text[start..end],
                    code, args, m_settings?.ErrorMessages, m_settings?.UiCulture)));
            }
        }

        /// <param name="settings">Determines the language of the error messages; <c>null</c> for English.</param>
        public static TemplateSyntaxTree Parse(string text, string part, ProcessSettings settings)
        {
            text ??= string.Empty;
            var ctx = new ParseContext(text, part, settings);

            // all patterns in document order - valid ones and those the pattern matcher rejected
            var found = new List<(int Index, int Length, PatternMatch Match, TemplateErrorCode Code, object[] Args)>();
            try
            {
                foreach (var m in PatternMatcher.FindSyntaxPatterns(text, (m, code, args) => found.Add((m.Index, m.Length, null, code, args))))
                {
                    found.Add((m.Index, m.Length, m, TemplateErrorCode.None, null));
                }
            }
            catch (OpenXmlTemplateException e)
            {
                ctx.AddError(0, 0, e.ErrorCode, e.Arguments.ToArray());
                return new TemplateSyntaxTree([], [], ctx.SortedErrors, settings);
            }
            found.Sort((a, b) => a.Index.CompareTo(b.Index));

            // Everything between {{:ignore}} and {{/:ignore}} is opaque text: no match, no error, no marker.
            var matches = new List<PatternMatch>();
            var covered = new List<(int Start, int End)>(); // text ranges consumed by a pattern
            bool inIgnore = false;
            foreach (var (index, length, match, code, args) in found)
            {
                if (inIgnore && match?.Type != PatternType.IgnoreEnd)
                {
                    continue;
                }
                covered.Add((index, index + length));
                if (match == null)
                {
                    ctx.AddError(index, length, code, args);
                    continue;
                }
                matches.Add(match);
                inIgnore = match.Type == PatternType.IgnoreBlock;
            }

            var root = BuildTree(matches, ctx);
            CheckUnrecognizedTags(text, covered, GetIgnoredRanges(root, text.Length), ctx);
            return new TemplateSyntaxTree(matches, root.Children, ctx.SortedErrors, settings);
        }

        /// <summary>
        /// Builds the block tree. Tags that are misplaced are reported and skipped, so the remaining structure can still
        /// be checked and all errors of a template are reported at once.
        /// </summary>
        private static TemplateSyntaxNode BuildTree(List<PatternMatch> matches, ParseContext ctx)
        {
            var root = new TemplateSyntaxNode(PatternType.None, null, null);
            var current = root; // always a content (None), else or separator node - never an opening block itself

            void Open(PatternMatch m)
            {
                var block = new TemplateSyntaxNode(m.Type, m, current);
                current = new TemplateSyntaxNode(PatternType.None, m, block);
            }

            void Close(PatternMatch m)
            {
                current.End = m;
                current.Parent.End = m;
                current = current.Parent.Parent;
            }

            foreach (var m in matches)
            {
                switch (m.Type)
                {
                    case PatternType.Expression:
                        CheckExpression(m, m.Variable, ctx);
                        break;
                    case PatternType.InlineKeyWord:
                        if (!InlineKeyWords.Contains(m.Variable.Trim()))
                        {
                            ctx.AddError(m, TemplateErrorCode.UnknownKeyword, m.Match.Value, string.Join(", ", InlineKeyWords.Select(x => $"{{{{:{x}}}}}")));
                        }
                        new TemplateSyntaxNode(m.Type, m, current).End = m;
                        break;
                    case PatternType.Condition:
                        CheckExpression(m, m.Condition, ctx);
                        Open(m);
                        break;
                    case PatternType.CollectionStart:
                        CheckCollectionName(m, ctx);
                        Open(m);
                        break;
                    case PatternType.Switch:
                        CheckExpression(m, GetKeywordArgument(m), ctx);
                        Open(m);
                        break;
                    case PatternType.RangeStart:
                    case PatternType.IgnoreBlock:
                        Open(m);
                        break;
                    case PatternType.Case:
                        CheckExpression(m, GetKeywordArgument(m), ctx);
                        goto case PatternType.Default;
                    case PatternType.Default:
                        // a case or default implicitly closes the preceding case / default
                        if (current.Parent?.Type is PatternType.Case or PatternType.Default)
                        {
                            Close(m);
                        }
                        if (!IsInside(current, PatternType.Switch))
                        {
                            ctx.AddError(m, TemplateErrorCode.CaseOutsideSwitch, m.Match.Value);
                        }
                        Open(m);
                        break;
                    case PatternType.ConditionElse:
                        if (current.Parent?.Type != PatternType.Condition)
                        {
                            ctx.AddError(m, TemplateErrorCode.ElseOutsideCondition, m.Match.Value);
                        }
                        else if (current.Type != PatternType.None)
                        {
                            ctx.AddError(m, TemplateErrorCode.DuplicateElse, current.Parent.Start.Match.Value, m.Match.Value);
                        }
                        else
                        {
                            current.End = m;
                            current = new TemplateSyntaxNode(m.Type, m, current.Parent);
                        }
                        break;
                    case PatternType.CollectionSeparator:
                        if (current.Parent?.Type != PatternType.CollectionStart || IsDynamicTable(current.Parent.Start))
                        {
                            ctx.AddError(m, TemplateErrorCode.SeparatorOutsideLoop, m.Match.Value);
                        }
                        else if (current.Type != PatternType.None)
                        {
                            ctx.AddError(m, TemplateErrorCode.DuplicateSeparator, current.Parent.Start.Match.Value, m.Match.Value);
                        }
                        else
                        {
                            current.End = m;
                            current = new TemplateSyntaxNode(m.Type, m, current.Parent);
                        }
                        break;
                    case PatternType.ConditionEnd:
                    case PatternType.CollectionEnd:
                    case PatternType.IgnoreEnd:
                        if (current == root)
                        {
                            ctx.AddError(m, TemplateErrorCode.NoMatchingOpeningTag, m.Match.Value);
                            break;
                        }
                        var closeError = CheckClosingTag(current.Parent.Start, m);
                        if (closeError is { } warning)
                        {
                            // rendering closes the current block regardless of the name, so this is only a warning
                            ctx.AddWarning(m, warning.Code, warning.Args);
                        }
                        Close(m);
                        break;
                }
            }

            for (var node = current; node != root; node = node.Parent.Parent)
            {
                ctx.AddError(node.Parent.Start, TemplateErrorCode.BlockNotClosed, node.Parent.Start.Match.Value);
            }
            return root;
        }

        private static bool IsInside(TemplateSyntaxNode node, PatternType type)
        {
            for (; node != null; node = node.Parent)
            {
                if (node.Type == type)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// <c>{{#switch}}</c> / <c>{{#case}}</c> without argument are parsed as loops over a collection named "switch" /
        /// "case" - almost certainly not intended. The short forms <c>{{#s}}</c> / <c>{{#c}}</c> could be real collection
        /// names, so they only get a warning.
        /// </summary>
        private static void CheckCollectionName(PatternMatch m, ParseContext ctx)
        {
            var name = m.Variable.Trim();
            if (name.Length == 0)
            {
                ctx.AddError(m, TemplateErrorCode.CollectionNameRequired, m.Match.Value);
            }
            else if (name.Equals("switch", StringComparison.OrdinalIgnoreCase) || name.Equals("case", StringComparison.OrdinalIgnoreCase))
            {
                ctx.AddError(m, TemplateErrorCode.SwitchExpressionRequired, m.Match.Value, name);
            }
            else if (name.Equals("s", StringComparison.OrdinalIgnoreCase) || name.Equals("c", StringComparison.OrdinalIgnoreCase))
            {
                ctx.AddWarning(m, TemplateErrorCode.SwitchLooksLikeLoop, m.Match.Value, name);
            }
        }

        /// <summary>The expression after the keyword of a switch / case tag, e.g. "ds.Val" for <c>{{#switch: ds.Val}}</c>.</summary>
        private static string GetKeywordArgument(PatternMatch m)
        {
            var colon = m.Variable.IndexOf(':');
            return colon < 0 ? string.Empty : m.Variable[(colon + 1)..];
        }

        private static (TemplateErrorCode Code, object[] Args)? CheckClosingTag(PatternMatch open, PatternMatch close)
        {
            var openTag = open.Match.Value;
            var closeTag = close.Match.Value;
            // an ignore block is always closed by its own end tag - everything else inside is skipped by Parse
            if (close.Type == PatternType.IgnoreEnd)
            {
                return open.Type == PatternType.IgnoreBlock ? null : (TemplateErrorCode.ClosingTagMismatch, [closeTag, openTag]);
            }
            if (close.Type == PatternType.CollectionEnd)
            {
                if (open.Type == PatternType.CollectionStart)
                {
                    return IsSameCollection(open.Variable, close.Variable) ? null : (TemplateErrorCode.ClosingTagMismatch, [closeTag, openTag]);
                }
                if (open.Type != PatternType.RangeStart)
                {
                    return (TemplateErrorCode.ClosingTagExpectedGeneric, [closeTag, openTag]);
                }
            }
            return null;
        }

        /// <summary>
        /// The closing tag may omit the implicit model prefix or leading dots, e.g. <c>{{#ds.Items}}...{{/Items}}</c>
        /// (the same way <c>{{Items.Value}}</c> resolves to <c>ds.Items.Value</c> at runtime).
        /// </summary>
        private static bool IsSameCollection(string openName, string closeName)
        {
            openName = openName.Trim().TrimStart('.');
            closeName = closeName.Trim().TrimStart('.');
            return string.Equals(openName, closeName, StringComparison.OrdinalIgnoreCase)
                || openName.EndsWith("." + closeName, StringComparison.OrdinalIgnoreCase)
                || closeName.EndsWith("." + openName, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsDynamicTable(PatternMatch match)
        {
            return string.Equals(match.Formatter, "dyntable", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Lightweight check of a condition or expression: not empty, string literals terminated and parentheses balanced.
        /// Whether the expression compiles can only be checked with a bound model, so these are warnings.
        /// </summary>
        private static void CheckExpression(PatternMatch m, string expression, ParseContext ctx)
        {
            if (string.IsNullOrWhiteSpace(expression))
            {
                ctx.AddWarning(m, TemplateErrorCode.EmptyExpression, m.Match.Value);
                return;
            }
            var sanitized = HelperFunctions.SanitizeQuotes(expression);
            var brackets = new Stack<char>();
            bool inString = false;
            for (int i = 0; i < sanitized.Length; i++)
            {
                var c = sanitized[i];
                if (inString)
                {
                    if (c == '\\')
                    {
                        i++;
                    }
                    else if (c == '"')
                    {
                        inString = false;
                    }
                    continue;
                }
                switch (c)
                {
                    case '"':
                        inString = true;
                        break;
                    case '(':
                    case '[':
                        brackets.Push(c);
                        break;
                    case ')':
                    case ']':
                        var expected = c == ')' ? '(' : '[';
                        if (brackets.Count == 0 || brackets.Pop() != expected)
                        {
                            ctx.AddWarning(m, TemplateErrorCode.UnbalancedBracket, c, expression);
                            return;
                        }
                        break;
                }
            }
            if (inString)
            {
                ctx.AddWarning(m, TemplateErrorCode.UnterminatedStringLiteral, expression);
            }
            else if (brackets.Count > 0)
            {
                ctx.AddWarning(m, TemplateErrorCode.UnclosedBracket, brackets.Peek(), expression);
            }
        }

        private static List<(int Start, int End)> GetIgnoredRanges(TemplateSyntaxNode root, int textLength)
        {
            var ranges = new List<(int Start, int End)>();
            Collect(root);
            return ranges;

            void Collect(TemplateSyntaxNode node)
            {
                foreach (var child in node.Children)
                {
                    if (child.Type == PatternType.IgnoreBlock)
                    {
                        ranges.Add((child.Start.Index, child.End != null ? child.End.Index + child.End.Length : textLength));
                    }
                    else
                    {
                        Collect(child);
                    }
                }
            }
        }

        /// <summary>
        /// Reports '{{' / '{?{' / '}}' outside of recognized tags - these are tags the pattern matcher could not
        /// parse (e.g. "{{Name}", "{{first name}}") and are left as plain text in the generated document.
        /// Content of :ignore blocks is skipped.
        /// </summary>
        private static void CheckUnrecognizedTags(string text, List<(int Start, int End)> covered, List<(int Start, int End)> ignoredRanges, ParseContext ctx)
        {
            int gapStart = 0;
            foreach (var (start, end) in covered.Append((text.Length, text.Length)))
            {
                if (start > gapStart && !ignoredRanges.Any(r => r.Start <= gapStart && start <= r.End))
                {
                    CheckGap(text, gapStart, start, ctx);
                }
                gapStart = Math.Max(gapStart, end);
            }
        }

        private static void CheckGap(string text, int gapStart, int gapEnd, ParseContext ctx)
        {
            var braces = StrayBracesRegex.Matches(text[gapStart..gapEnd]).Cast<Match>().ToList();
            for (int i = 0; i < braces.Count; i++)
            {
                var brace = braces[i];
                var index = gapStart + brace.Index;
                if (brace.Groups["close"].Success)
                {
                    ctx.AddWarning(index, brace.Length, TemplateErrorCode.UnexpectedClosingBraces, brace.Value);
                    continue;
                }
                var next = i + 1 < braces.Count ? braces[i + 1] : null;
                if (next != null && next.Groups["close"].Success)
                {
                    var length = next.Index + next.Length - brace.Index;
                    ctx.AddWarning(index, length, TemplateErrorCode.InvalidTag, text.Substring(index, length));
                    i++;
                }
                else
                {
                    var end = next != null ? gapStart + next.Index : gapEnd;
                    var length = Math.Min(MaxUnterminatedTagLength, end - index);
                    ctx.AddWarning(index, length, TemplateErrorCode.TagNotTerminated, text.Substring(index, length));
                }
            }
        }
    }
}
