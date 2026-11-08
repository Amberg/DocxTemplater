using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DocxTemplater.Text
{
    /// <summary>
    /// Navigation in the <see cref="TemplateSyntaxTree"/> of a text template, shared by the renderer, the schema
    /// collector and the validation. The tree has the shape of the block tree of the Word renderer: a block node
    /// (loop, condition, ...) has one content node per section (body, else / separator), and a section's text
    /// is what lies between its start tag and its end tag.
    /// </summary>
    internal sealed class SyntaxTreeHelper
    {
        private readonly PatternMatch[] m_values;

        public SyntaxTreeHelper(TemplateSyntaxTree tree)
        {
            m_values = tree.Matches.Where(x => x.Type is PatternType.Variable or PatternType.Expression).ToArray();
        }

        /// <summary>The value tags (variables and expressions) in <c>[from, to)</c>, in document order.</summary>
        public IEnumerable<PatternMatch> ValuesIn(int from, int to)
        {
            int low = 0, high = m_values.Length;
            while (low < high)
            {
                int mid = (low + high) / 2;
                if (m_values[mid].Index < from)
                {
                    low = mid + 1;
                }
                else
                {
                    high = mid;
                }
            }
            for (int i = low; i < m_values.Length && m_values[i].Index < to; i++)
            {
                yield return m_values[i];
            }
        }

        /// <summary>Start of the text of a section (content, else or separator node): right after its start tag.</summary>
        public static int SectionStart(TemplateSyntaxNode section)
        {
            return section.Start.Index + section.Start.Length;
        }

        /// <summary>End of the text of a section: the start of its end tag (else, separator, case or closing tag).</summary>
        public static int SectionEnd(TemplateSyntaxNode section)
        {
            return section.End.Index;
        }

        /// <summary>First position after a node (after its closing tag; an inline tag is its own end).</summary>
        public static int NodeEnd(TemplateSyntaxNode node)
        {
            // a case that is implicitly closed by the next case / default does not own its end tag - the next node starts with it
            return node.End.Type is PatternType.Case or PatternType.Default ? node.End.Index : node.End.Index + node.End.Length;
        }

        /// <summary>The expression after the keyword of a switch / case tag, e.g. "ds.Val" for <c>{{#switch: ds.Val}}</c>.</summary>
        public static string KeywordArgument(PatternMatch match)
        {
            var variable = match.Variable ?? string.Empty;
            var colon = variable.IndexOf(':');
            return colon < 0 ? string.Empty : variable[(colon + 1)..].Trim();
        }

        /// <summary>A quoted string or a number: a case expression that references no model path.</summary>
        public static bool IsLiteral(string expression)
        {
            if (expression.Length == 0 || expression[0] is '\'' or '"')
            {
                return true;
            }
            return double.TryParse(expression, NumberStyles.Number, CultureInfo.InvariantCulture, out _);
        }

        /// <summary>
        /// True for the tags that structure the template - everything but values and inline keywords. Only lines
        /// made of these are removed by <see cref="TextTemplateSettings.IgnoreLineBreaksAroundTags"/>.
        /// </summary>
        public static bool IsBlockTag(PatternMatch match)
        {
            return match.Type is not (PatternType.Variable or PatternType.Expression or PatternType.InlineKeyWord);
        }

        /// <summary>
        /// The ranges of the template text that <see cref="TextTemplateSettings.IgnoreLineBreaksAroundTags"/>
        /// removes: for every line that holds nothing but block tags and whitespace, the indentation before the
        /// first tag and everything after the last tag including the line break. Whitespace between two tags
        /// stays - it is the content of the block between them. Sorted and non-overlapping.
        /// </summary>
        public static IReadOnlyList<(int Start, int End)> LineBreakRangesToSkip(string text, IReadOnlyList<PatternMatch> matches)
        {
            var result = new List<(int Start, int End)>();
            var blockTags = matches.Where(IsBlockTag).OrderBy(x => x.Index).ToList();
            int tag = 0;
            int lineStart = 0;
            while (lineStart < text.Length)
            {
                int newline = text.IndexOf('\n', lineStart);
                int nextLineStart = newline < 0 ? text.Length : newline + 1;
                int lineEnd = newline < 0 ? text.Length : newline;
                if (lineEnd > lineStart && text[lineEnd - 1] == '\r')
                {
                    lineEnd--;
                }

                var tagsOnLine = new List<PatternMatch>();
                for (; tag < blockTags.Count && blockTags[tag].Index < nextLineStart; tag++)
                {
                    if (blockTags[tag].Index >= lineStart && blockTags[tag].Index + blockTags[tag].Length <= lineEnd)
                    {
                        tagsOnLine.Add(blockTags[tag]);
                    }
                }

                if (tagsOnLine.Count > 0 && OnlyWhitespaceBetween(text, lineStart, lineEnd, tagsOnLine))
                {
                    AddRange(result, lineStart, tagsOnLine[0].Index);
                    var last = tagsOnLine[^1];
                    AddRange(result, last.Index + last.Length, nextLineStart);
                }
                lineStart = nextLineStart;
            }
            return result;
        }

        private static bool OnlyWhitespaceBetween(string text, int lineStart, int lineEnd, List<PatternMatch> tags)
        {
            int pos = lineStart;
            foreach (var tag in tags)
            {
                if (!IsWhitespace(text, pos, tag.Index))
                {
                    return false;
                }
                pos = tag.Index + tag.Length;
            }
            return IsWhitespace(text, pos, lineEnd);
        }

        private static bool IsWhitespace(string text, int from, int to)
        {
            for (int i = from; i < to; i++)
            {
                if (!char.IsWhiteSpace(text[i]))
                {
                    return false;
                }
            }
            return true;
        }

        private static void AddRange(List<(int Start, int End)> ranges, int start, int end)
        {
            if (end > start)
            {
                ranges.Add((start, end));
            }
        }
    }
}
