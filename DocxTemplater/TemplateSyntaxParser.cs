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
        public TemplateSyntaxTree(IReadOnlyList<PatternMatch> matches, IReadOnlyList<TemplateSyntaxNode> blocks, IReadOnlyList<TemplateSyntaxError> errors)
        {
            Matches = matches;
            Blocks = blocks;
            Errors = errors;
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
                var errors = Errors.Where(x => x.Severity == TemplateSyntaxErrorSeverity.Error);
                throw new OpenXmlTemplateException($"Template syntax errors:{Environment.NewLine}{string.Join(Environment.NewLine, errors)}");
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
            private readonly List<(int Index, TemplateSyntaxError Error)> m_errors = [];

            public ParseContext(string text, string part)
            {
                m_text = text;
                m_part = part;
            }

            public void AddError(int index, int length, string message)
            {
                Add(TemplateSyntaxErrorSeverity.Error, index, length, message);
            }

            public void AddError(PatternMatch match, string message)
            {
                Add(TemplateSyntaxErrorSeverity.Error, match.Index, match.Length, message);
            }

            public void AddWarning(PatternMatch match, string message)
            {
                Add(TemplateSyntaxErrorSeverity.Warning, match.Index, match.Length, message);
            }

            public void AddWarning(int index, int length, string message)
            {
                Add(TemplateSyntaxErrorSeverity.Warning, index, length, message);
            }

            public IReadOnlyList<TemplateSyntaxError> SortedErrors => m_errors.OrderBy(x => x.Index).Select(x => x.Error).ToList();

            private void Add(TemplateSyntaxErrorSeverity severity, int index, int length, string message)
            {
                var start = Math.Max(0, index - ContextLength);
                var end = Math.Min(m_text.Length, index + length + ContextLength);
                m_errors.Add((index, new TemplateSyntaxError(severity, m_part, m_text.Substring(index, length), message, m_text[start..end])));
            }
        }

        public static TemplateSyntaxTree Parse(string text, string part)
        {
            text ??= string.Empty;
            var ctx = new ParseContext(text, part);

            // text ranges consumed by a pattern - valid or invalid
            var covered = new List<(int Start, int End)>();
            List<PatternMatch> matches;
            try
            {
                matches = PatternMatcher.FindSyntaxPatterns(text, (m, message) =>
                {
                    covered.Add((m.Index, m.Index + m.Length));
                    ctx.AddError(m.Index, m.Length, message);
                }).ToList();
            }
            catch (OpenXmlTemplateException e)
            {
                ctx.AddError(0, 0, e.Message);
                return new TemplateSyntaxTree([], [], ctx.SortedErrors);
            }
            covered.AddRange(matches.Select(m => (m.Index, m.Index + m.Length)));
            covered.Sort((a, b) => a.Start.CompareTo(b.Start));

            var root = BuildTree(matches, ctx);
            CheckUnrecognizedTags(text, covered, GetIgnoredRanges(root, text.Length), ctx);
            return new TemplateSyntaxTree(matches, root.Children, ctx.SortedErrors);
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
                            ctx.AddError(m, $"Unknown keyword '{m.Match.Value}'. Supported keywords are {string.Join(", ", InlineKeyWords.Select(x => $"{{{{:{x}}}}}"))}");
                        }
                        new TemplateSyntaxNode(m.Type, m, current).End = m;
                        break;
                    case PatternType.Condition:
                        CheckExpression(m, m.Condition, ctx);
                        Open(m);
                        break;
                    case PatternType.CollectionStart:
                        var name = m.Variable.Trim();
                        if (name.Equals("switch", StringComparison.OrdinalIgnoreCase) || name.Equals("case", StringComparison.OrdinalIgnoreCase))
                        {
                            ctx.AddError(m, $"'{m.Match.Value}' requires an expression, e.g. '{{{{#{name}: Value}}}}'");
                        }
                        Open(m);
                        break;
                    case PatternType.Switch:
                    case PatternType.RangeStart:
                    case PatternType.IgnoreBlock:
                        Open(m);
                        break;
                    case PatternType.Case:
                    case PatternType.Default:
                        // a case or default implicitly closes the preceding case / default
                        if (current.Parent?.Type is PatternType.Case or PatternType.Default)
                        {
                            Close(m);
                        }
                        if (!IsInside(current, PatternType.Switch))
                        {
                            ctx.AddError(m, $"'{m.Match.Value}' must be inside a '{{{{#switch: ...}}}}' block");
                        }
                        Open(m);
                        break;
                    case PatternType.ConditionElse:
                        if (current.Parent?.Type != PatternType.Condition)
                        {
                            ctx.AddError(m, $"'{m.Match.Value}' (else) is only allowed directly inside a condition '{{?{{...}}}}'");
                        }
                        else if (current.Type != PatternType.None)
                        {
                            ctx.AddError(m, $"Condition '{current.Parent.Start.Match.Value}' has more than one else '{m.Match.Value}'");
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
                            ctx.AddError(m, $"Separator '{m.Match.Value}' is only allowed directly inside a collection loop '{{{{#Items}}}}'");
                        }
                        else if (current.Type != PatternType.None)
                        {
                            ctx.AddError(m, $"Loop '{current.Parent.Start.Match.Value}' has more than one separator '{m.Match.Value}'");
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
                            ctx.AddError(m, $"'{m.Match.Value}' has no matching opening tag");
                            break;
                        }
                        var closeError = CheckClosingTag(current.Parent.Start, m);
                        if (closeError != null)
                        {
                            ctx.AddError(m, closeError);
                        }
                        Close(m);
                        break;
                }
            }

            for (var node = current; node != root; node = node.Parent.Parent)
            {
                ctx.AddError(node.Parent.Start, $"'{node.Parent.Start.Match.Value}' is not closed");
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

        private static string CheckClosingTag(PatternMatch open, PatternMatch close)
        {
            var openTag = open.Match.Value;
            var closeTag = close.Match.Value;
            if (open.Type == PatternType.IgnoreBlock)
            {
                return close.Type == PatternType.IgnoreEnd ? null : $"'{closeTag}' closes '{openTag}', expected '{{{{/:ignore}}}}'";
            }
            if (close.Type == PatternType.IgnoreEnd)
            {
                return $"'{closeTag}' does not match '{openTag}'";
            }
            if (close.Type == PatternType.CollectionEnd)
            {
                if (open.Type == PatternType.CollectionStart)
                {
                    return IsSameCollection(open.Variable, close.Variable) ? null : $"'{closeTag}' does not match '{openTag}'";
                }
                if (open.Type != PatternType.RangeStart)
                {
                    return $"'{closeTag}' closes '{openTag}', expected '{{{{/}}}}'";
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
                ctx.AddWarning(m, $"'{m.Match.Value}' has an empty expression");
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
                            ctx.AddWarning(m, $"Unbalanced '{c}' in expression '{expression}'");
                            return;
                        }
                        break;
                }
            }
            if (inString)
            {
                ctx.AddWarning(m, $"Unterminated string literal in expression '{expression}'");
            }
            else if (brackets.Count > 0)
            {
                ctx.AddWarning(m, $"Unclosed '{brackets.Peek()}' in expression '{expression}'");
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
                    ctx.AddWarning(index, brace.Length, $"Unexpected '{brace.Value}' without matching '{{{{'");
                    continue;
                }
                var next = i + 1 < braces.Count ? braces[i + 1] : null;
                if (next != null && next.Groups["close"].Success)
                {
                    var length = next.Index + next.Length - brace.Index;
                    ctx.AddWarning(index, length, $"Invalid tag '{text.Substring(index, length)}'");
                    i++;
                }
                else
                {
                    var end = next != null ? gapStart + next.Index : gapEnd;
                    var length = Math.Min(MaxUnterminatedTagLength, end - index);
                    ctx.AddWarning(index, length, $"Tag '{text.Substring(index, length)}' is not terminated");
                }
            }
        }
    }
}
