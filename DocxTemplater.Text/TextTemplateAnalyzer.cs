using System;
using System.Collections.Generic;
using System.Globalization;
using DocxTemplater.Schema;

namespace DocxTemplater.Text
{
    /// <summary>
    /// The static analysis of a text template: the constructs that exist only in Word templates, and the
    /// <see cref="TemplateSchema"/> - the same walk over the syntax tree as the renderer, declaring paths instead of
    /// resolving them, with the scope rules of the blocks of the Word renderer.
    /// </summary>
    internal sealed class TextTemplateAnalyzer
    {
        private const int ContextLength = 30;

        private static readonly HashSet<string> WordOnlyFormatters = new(StringComparer.OrdinalIgnoreCase) { "img", "image", "template", "t", "md", "dyntable" };
        private static readonly HashSet<string> WordOnlyKeywords = new(StringComparer.OrdinalIgnoreCase) { "PageBreak", "SectionBreak" };

        private readonly SyntaxTreeHelper m_helper;
        private readonly SchemaBuilder m_builder = new();

        private TextTemplateAnalyzer(TemplateSyntaxTree tree)
        {
            m_helper = new SyntaxTreeHelper(tree);
        }

        /// <summary>
        /// Errors for tags a text template cannot render: formatters that insert Word content (images, sub-templates,
        /// markdown, dynamic tables) and the page / section break keywords.
        /// </summary>
        public static IEnumerable<TemplateSyntaxError> FindWordOnlyConstructs(TemplateSyntaxTree tree, string text, string part, ProcessSettings settings)
        {
            foreach (var match in tree.Matches)
            {
                if (match.Type == PatternType.InlineKeyWord && WordOnlyKeywords.Contains(match.Variable.Trim()))
                {
                    yield return CreateError(match, text, part, settings, TemplateErrorCode.KeywordNotSupported, match.Match.Value);
                }
                else if (!string.IsNullOrEmpty(match.Formatter) && WordOnlyFormatters.Contains(match.Formatter))
                {
                    yield return CreateError(match, text, part, settings, TemplateErrorCode.FormatterNotSupported, match.Match.Value, match.Formatter);
                }
            }
        }

        private static TemplateSyntaxError CreateError(PatternMatch match, string text, string part, ProcessSettings settings, TemplateErrorCode code, params object[] arguments)
        {
            var start = Math.Max(0, match.Index - ContextLength);
            var end = Math.Min(text.Length, match.Index + match.Length + ContextLength);
            return new TemplateSyntaxError(TemplateSyntaxErrorSeverity.Error, part, match.Match.Value, text[start..end],
                code, arguments, settings.ErrorMessages, settings.UiCulture, match.Index, match.Length);
        }

        public static TemplateSchema CollectSchema(TemplateSyntaxTree tree, int textLength)
        {
            var analyzer = new TextTemplateAnalyzer(tree);
            analyzer.CollectContent(tree.Blocks, 0, textLength);
            return analyzer.m_builder.Build();
        }

        private void CollectContent(IReadOnlyList<TemplateSyntaxNode> children, int from, int to)
        {
            int pos = from;
            foreach (var child in children)
            {
                DeclareValues(pos, child.Start.Index);
                CollectNode(child);
                pos = SyntaxTreeHelper.NodeEnd(child);
            }
            DeclareValues(pos, to);
        }

        private void DeclareValues(int from, int to)
        {
            foreach (var value in m_helper.ValuesIn(from, to))
            {
                if (value.Type == PatternType.Variable)
                {
                    m_builder.DeclareScalar(value.Variable);
                }
                else
                {
                    // the expression includes the outer parentheses, which the parser treats like the Word renderer does
                    SchemaExpressionParser.Extract(value.Variable, m_builder.DeclareScalar);
                }
            }
        }

        private void CollectSections(TemplateSyntaxNode node)
        {
            foreach (var section in node.Children)
            {
                CollectContent(section.Children, SyntaxTreeHelper.SectionStart(section), SyntaxTreeHelper.SectionEnd(section));
            }
        }

        private void CollectNode(TemplateSyntaxNode node)
        {
            switch (node.Type)
            {
                case PatternType.CollectionStart:
                    CollectLoop(node);
                    break;
                case PatternType.Condition:
                    SchemaExpressionParser.Extract(node.Start.Condition, m_builder.DeclareScalar);
                    CollectSections(node);
                    break;
                case PatternType.Switch:
                    SchemaExpressionParser.Extract(SyntaxTreeHelper.KeywordArgument(node.Start), m_builder.DeclareScalar);
                    CollectSections(node);
                    break;
                case PatternType.Case:
                    var caseExpression = SyntaxTreeHelper.KeywordArgument(node.Start);
                    if (!SyntaxTreeHelper.IsLiteral(caseExpression))
                    {
                        SchemaExpressionParser.Extract(caseExpression, m_builder.DeclareScalar);
                    }
                    CollectSections(node);
                    break;
                case PatternType.Default:
                    CollectSections(node);
                    break;
                case PatternType.RangeStart:
                    CollectRangeLoop(node);
                    break;
                default:
                    // ignore blocks and inline keywords reference no model paths
                    break;
            }
        }

        private void CollectLoop(TemplateSyntaxNode node)
        {
            var collectionName = node.Start.Variable;
            // the name can be dot-prefixed (a nested loop on the outer item); the builder needs the absolute key of the scope
            var absoluteKey = m_builder.ResolveAbsoluteScopeKey(collectionName);
            var itemSchema = m_builder.DeclareCollection(collectionName);
            if (itemSchema != null)
            {
                m_builder.OpenItemScope(absoluteKey, itemSchema);
            }
            else
            {
                m_builder.OpenTransparentScope();
            }
            try
            {
                CollectSections(node);
            }
            finally
            {
                m_builder.CloseScope();
            }
        }

        private void CollectRangeLoop(TemplateSyntaxNode node)
        {
            var (_, count) = RangeLoopCount.Parse(node.Start.Variable);
            if (!int.TryParse(count, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
            {
                SchemaExpressionParser.Extract(count, m_builder.DeclareScalar);
            }
            // a range loop introduces a scope level for dot-resolution but no named item
            m_builder.OpenTransparentScope();
            try
            {
                CollectSections(node);
            }
            finally
            {
                m_builder.CloseScope();
            }
        }
    }
}
