using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DocxTemplater
{
    public partial class TemplateErrorMessages
    {
        /// <summary>
        /// The built-in English texts. Use them as a template for your own language dictionary.
        /// </summary>
        public static IReadOnlyDictionary<TemplateErrorCode, string> English { get; } = new ReadOnlyDictionary<TemplateErrorCode, string>(
            new Dictionary<TemplateErrorCode, string>
            {
                // model binding
                [TemplateErrorCode.ModelPrefixAlreadyBound] = "A model with the prefix '{0}' has already been bound - model prefixes are case-insensitive",
                [TemplateErrorCode.ModelNotFound] = "Model {0} not found",
                [TemplateErrorCode.PropertyNotFound] = "Property {0} not found in {1}",
                [TemplateErrorCode.PropertyNotFoundInParentScope] = "Property {0} not found in parent scope",
                [TemplateErrorCode.PropertyNotFoundOnCollection] = "Property '{0}' on collection of type '{1}' not found",
                [TemplateErrorCode.PropertyNotFoundOnType] = "Property '{0}' not found in '{1}' of type '{2}'",
                [TemplateErrorCode.NotEnumerable] = "'{0}' is not enumerable - it is of type {1}",
                [TemplateErrorCode.RangeCountNotNumeric] = "'{0}' is not an integer - its value '{1}' cannot be parsed to an integer",
                [TemplateErrorCode.RangeCountNotNumericOrEnumerable] = "'{0}' is not an integer or enumerable - it is of type {1}",
                [TemplateErrorCode.NotADynamicTable] = "'{0}' is not of type {1}",
                [TemplateErrorCode.IndexingNotSupported] = "Cannot apply indexing to an expression of type '{0}'",

                // expressions
                [TemplateErrorCode.ScriptParseError] = "Error parsing script {0}",
                [TemplateErrorCode.ExpressionParseError] = "Error parsing expression {0}",
                [TemplateErrorCode.InvalidExpression] = "Invalid expression '{0}'",
                [TemplateErrorCode.ErrorInCondition] = "{0} in condition '{1}'",
                [TemplateErrorCode.ErrorInSwitchCase] = "{0} in switch '{1}' case '{2}'",

                // placeholder replacement
                [TemplateErrorCode.InvalidVariableSyntax] = "Invalid variable syntax '{0}'",
                [TemplateErrorCode.PlaceholderNotReplaced] = "'{0}' could not be replaced: {1} >> {0} << {2}",
                [TemplateErrorCode.ContentControlNotReplaced] = "Content control tag '{0}' could not be replaced: {1}",
                [TemplateErrorCode.ContentControlHasNoText] = "the resolved value has no text run to fill (empty cell/row content controls are not supported)",

                // formatters
                [TemplateErrorCode.FormatPatternRequiresArgument] = "DateTime formatter requires exactly one argument, e.g. FORMAT(dd.MM.yyyy)",
                [TemplateErrorCode.FormatNotApplicable] = "Format {0} cannot be applied to {1} of type {2}",
                [TemplateErrorCode.FormatterRequiresFormattable] = "Formatter {0} can only be applied to IFormattable objects - property {1}",
                [TemplateErrorCode.FormatterRequiresString] = "Formatter {0} can only be applied to string objects - property {1}",
                [TemplateErrorCode.InvalidFormatterArguments] = "Arguments must be in the form foo:bar",
                [TemplateErrorCode.SubTemplateNameRequired] = "Template formatter requires a template name",
                [TemplateErrorCode.SubTemplateInvalid] = "Template is null or is not a valid OpenXML template",
                [TemplateErrorCode.SubTemplateInvalidArgument] = "Invalid template formatter argument '{0}'",
                [TemplateErrorCode.SubTemplateInvalidSelector] = "Invalid selector {0}",
                [TemplateErrorCode.SubTemplateUnsupportedElement] = "Template must be a paragraph, run, table row or table cell",
                [TemplateErrorCode.SubTemplateUnsupportedType] = "Template must be a string, OpenXmlElement, byte[] or Stream",
                [TemplateErrorCode.HtmlNotInParagraph] = "HTML import tag is not in a paragraph",
                [TemplateErrorCode.ImageMetadataUnreadable] = "Could not read image metadata",
                [TemplateErrorCode.InvalidImageFormatterArgument] = "Invalid image formatter argument '{0}'",
                [TemplateErrorCode.MarkdownNestingDepthExceeded] = "Markdown nesting depth exceeded",
                [TemplateErrorCode.MarkdownInvalidImage] = "Invalid image data in Markdown link. {0}",

                // template syntax
                [TemplateErrorCode.TemplateSyntaxErrors] = "Template syntax errors:\n{0}",
                [TemplateErrorCode.SyntaxErrorLocation] = "{0}: {1} (near '{2}')",
                [TemplateErrorCode.InvalidSyntax] = "Invalid syntax '{0}'",
                [TemplateErrorCode.UseGenericClosingTag] = "Invalid syntax '{0}'. Use '{{{{/}}}}' instead.",
                [TemplateErrorCode.RangeLoopVariableNameRequired] = "Invalid range loop syntax '{0}' - variable name is required",
                [TemplateErrorCode.PatternMatchTimeout] = "Invalid syntax '{0}' - match timeout",
                [TemplateErrorCode.UnknownKeyword] = "Unknown keyword '{0}'. Supported keywords are {1}",
                [TemplateErrorCode.CaseOutsideSwitch] = "'{0}' must be inside a '{{{{#switch: ...}}}}' block",
                [TemplateErrorCode.ElseOutsideCondition] = "'{0}' (else) is only allowed directly inside a condition '{{?{{...}}}}'",
                [TemplateErrorCode.DuplicateElse] = "Condition '{0}' has more than one else '{1}'",
                [TemplateErrorCode.SeparatorOutsideLoop] = "Separator '{0}' is only allowed directly inside a collection loop '{{{{#Items}}}}'",
                [TemplateErrorCode.DuplicateSeparator] = "Loop '{0}' has more than one separator '{1}'",
                [TemplateErrorCode.NoMatchingOpeningTag] = "'{0}' has no matching opening tag",
                [TemplateErrorCode.BlockNotClosed] = "'{0}' is not closed",
                [TemplateErrorCode.CollectionNameRequired] = "'{0}' requires a collection name, e.g. '{{{{#Items}}}}'",
                [TemplateErrorCode.SwitchExpressionRequired] = "'{0}' requires an expression, e.g. '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.SwitchLooksLikeLoop] = "'{0}' is a loop over a collection named '{1}'. For a switch / case use '{{{{#{1}: Value}}}}'",
                [TemplateErrorCode.ClosingTagMismatch] = "'{0}' does not match '{1}'",
                [TemplateErrorCode.ClosingTagExpectedGeneric] = "'{0}' closes '{1}', expected '{{{{/}}}}'",
                [TemplateErrorCode.EmptyExpression] = "'{0}' has an empty expression",
                [TemplateErrorCode.UnbalancedBracket] = "Unbalanced '{0}' in expression '{1}'",
                [TemplateErrorCode.UnterminatedStringLiteral] = "Unterminated string literal in expression '{0}'",
                [TemplateErrorCode.UnclosedBracket] = "Unclosed '{0}' in expression '{1}'",
                [TemplateErrorCode.UnexpectedClosingBraces] = "Unexpected '{0}' without matching '{{{{'",
                [TemplateErrorCode.InvalidTag] = "Invalid tag '{0}'",
                [TemplateErrorCode.TagNotTerminated] = "Tag '{0}' is not terminated",
                [TemplateErrorCode.InvalidSwitchBlockSyntax] = "Invalid switch block syntax: {0}",
                [TemplateErrorCode.InvalidCaseBlockSyntax] = "Invalid case block syntax: {0}",
                [TemplateErrorCode.CaseNotNestedInSwitch] = "The '{{#case}}'/'{{#default}}' block ('{0}') must be nested inside a '{{#switch}}' block.",
                [TemplateErrorCode.InvalidInlineKeyword] = "Invalid expression {0}",
                [TemplateErrorCode.DynamicTableStructureInvalid] = "Dynamic table block must contain exactly one table with at least a header and a data row",
                [TemplateErrorCode.DynamicTableSingleChild] = "Dynamic table block must contain exactly one child block"
            });
    }
}
