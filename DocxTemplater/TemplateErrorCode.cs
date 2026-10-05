namespace DocxTemplater
{
    /// <summary>
    /// Identifies an error reported by the template engine in an <see cref="OpenXmlTemplateException"/> or a
    /// <see cref="TemplateSyntaxError"/>, independent of the language of its message. The documentation of each
    /// member lists the positional arguments (<c>{0}</c>, <c>{1}</c>, ...) available to the message format, see
    /// <see cref="TemplateErrorMessages"/>.
    /// </summary>
    public enum TemplateErrorCode
    {
        /// <summary>
        /// The message is free text that is not translated (internal errors, errors of third party extensions).
        /// </summary>
        None = 0,

        // ----- model binding -----

        /// <summary>A second model was bound with an already used prefix. {0}: prefix</summary>
        ModelPrefixAlreadyBound,

        /// <summary>No bound model or scope variable matches the placeholder. {0}: variable name</summary>
        ModelNotFound,

        /// <summary>A property of a model implementing <c>ITemplateModel</c> or a dictionary key does not exist. {0}: property name, {1}: model path</summary>
        PropertyNotFound,

        /// <summary>A property of the parent scope (<c>..Name</c>) does not exist. {0}: property name</summary>
        PropertyNotFoundInParentScope,

        /// <summary>A property was accessed on a collection. {0}: variable name, {1}: type of the collection</summary>
        PropertyNotFoundOnCollection,

        /// <summary>A CLR property does not exist on the model. {0}: property name, {1}: variable name, {2}: type of the model</summary>
        PropertyNotFoundOnType,

        /// <summary>A loop variable is not enumerable. {0}: variable name, {1}: type of the value</summary>
        NotEnumerable,

        /// <summary>The count of a range loop is a string that is not an integer. {0}: variable name, {1}: value</summary>
        RangeCountNotNumeric,

        /// <summary>The count of a range loop is neither an integer nor enumerable. {0}: variable name, {1}: type of the value</summary>
        RangeCountNotNumericOrEnumerable,

        /// <summary>The variable of a dynamic table does not implement <c>IDynamicTable</c>. {0}: variable name, {1}: expected type</summary>
        NotADynamicTable,

        /// <summary>An indexer <c>[i]</c> was applied to a value that does not support indexing. {0}: type of the value</summary>
        IndexingNotSupported,

        // ----- expressions -----

        /// <summary>A condition or case expression could not be parsed. {0}: script</summary>
        ScriptParseError,

        /// <summary>An expression placeholder <c>{{(...)}}</c> could not be parsed. {0}: expression</summary>
        ExpressionParseError,

        /// <summary>An expression could not be pre-processed. {0}: expression</summary>
        InvalidExpression,

        /// <summary>Wraps an error that occurred while evaluating a condition. {0}: inner message, {1}: condition</summary>
        ErrorInCondition,

        /// <summary>Wraps an error that occurred while evaluating a switch case. {0}: inner message, {1}: switch expression, {2}: case expression</summary>
        ErrorInSwitchCase,

        // ----- placeholder replacement -----

        /// <summary>The text of a placeholder could not be parsed. {0}: placeholder text</summary>
        InvalidVariableSyntax,

        /// <summary>Wraps an error that occurred while replacing a placeholder. {0}: placeholder text, {1}: text before, {2}: text after</summary>
        PlaceholderNotReplaced,

        /// <summary>Wraps an error that occurred while filling a content control. {0}: tag, {1}: inner message</summary>
        ContentControlNotReplaced,

        /// <summary>A content control has no text run that could hold the value.</summary>
        ContentControlHasNoText,

        // ----- formatters -----

        /// <summary>The FORMAT formatter was used without a format pattern.</summary>
        FormatPatternRequiresArgument,

        /// <summary>A format pattern does not apply to the value. {0}: format pattern, {1}: placeholder, {2}: type of the value</summary>
        FormatNotApplicable,

        /// <summary>The FORMAT formatter was applied to a value that is not <c>IFormattable</c>. {0}: formatter, {1}: placeholder</summary>
        FormatterRequiresFormattable,

        /// <summary>A string formatter (toupper / tolower) was applied to a non-string value. {0}: formatter, {1}: placeholder</summary>
        FormatterRequiresString,

        /// <summary>Formatter arguments are not in the form <c>name:value</c>.</summary>
        InvalidFormatterArguments,

        /// <summary>The template formatter was used without a template name.</summary>
        SubTemplateNameRequired,

        /// <summary>The value of a template formatter is null or not a valid OpenXML document.</summary>
        SubTemplateInvalid,

        /// <summary>An argument of the template formatter is not a known selector. {0}: argument</summary>
        SubTemplateInvalidArgument,

        /// <summary>The selector of a template formatter is unknown. {0}: selector</summary>
        SubTemplateInvalidSelector,

        /// <summary>The root element of a sub-template is not a paragraph, run, table row or table cell.</summary>
        SubTemplateUnsupportedElement,

        /// <summary>The value of a template formatter has an unsupported type.</summary>
        SubTemplateUnsupportedType,

        /// <summary>The HTML formatter was used outside of a paragraph.</summary>
        HtmlNotInParagraph,

        /// <summary>The image formatter could not read the metadata of the image.</summary>
        ImageMetadataUnreadable,

        /// <summary>An argument of the image formatter is invalid. {0}: argument</summary>
        InvalidImageFormatterArgument,

        /// <summary>Markdown formatters are nested too deeply.</summary>
        MarkdownNestingDepthExceeded,

        /// <summary>An image link in markdown does not contain valid image data. {0}: link</summary>
        MarkdownInvalidImage,

        // ----- template syntax -----

        /// <summary>Rendering was aborted because the template has syntax errors. {0}: list of errors, one per line</summary>
        TemplateSyntaxErrors,

        /// <summary>Location line of a <see cref="TemplateSyntaxError"/>. {0}: document part, {1}: message, {2}: surrounding text</summary>
        SyntaxErrorLocation,

        /// <summary>A tag could not be parsed. {0}: tag</summary>
        InvalidSyntax,

        /// <summary>A closing tag names a keyword (<c>{{/switch}}</c>). {0}: tag</summary>
        UseGenericClosingTag,

        /// <summary>A range loop tag without a variable name. {0}: tag</summary>
        RangeLoopVariableNameRequired,

        /// <summary>Parsing the template text timed out. {0}: text</summary>
        PatternMatchTimeout,

        /// <summary>An inline keyword is unknown. {0}: tag, {1}: list of supported keywords</summary>
        UnknownKeyword,

        /// <summary>A case or default tag outside of a switch block. {0}: tag</summary>
        CaseOutsideSwitch,

        /// <summary>An else tag outside of a condition. {0}: tag</summary>
        ElseOutsideCondition,

        /// <summary>A condition has more than one else tag. {0}: condition tag, {1}: else tag</summary>
        DuplicateElse,

        /// <summary>A separator tag outside of a collection loop. {0}: tag</summary>
        SeparatorOutsideLoop,

        /// <summary>A loop has more than one separator tag. {0}: loop tag, {1}: separator tag</summary>
        DuplicateSeparator,

        /// <summary>A closing tag without an opening tag. {0}: tag</summary>
        NoMatchingOpeningTag,

        /// <summary>A block is not closed. {0}: opening tag</summary>
        BlockNotClosed,

        /// <summary>A loop tag without a collection name. {0}: tag</summary>
        CollectionNameRequired,

        /// <summary>A switch or case tag without an expression. {0}: tag, {1}: keyword</summary>
        SwitchExpressionRequired,

        /// <summary>A loop over a collection named like a switch / case shortcut. {0}: tag, {1}: name</summary>
        SwitchLooksLikeLoop,

        /// <summary>A closing tag does not match its opening tag. {0}: closing tag, {1}: opening tag</summary>
        ClosingTagMismatch,

        /// <summary>A named closing tag closes a block that expects the generic closing tag. {0}: closing tag, {1}: opening tag</summary>
        ClosingTagExpectedGeneric,

        /// <summary>A tag has an empty expression. {0}: tag</summary>
        EmptyExpression,

        /// <summary>A closing bracket without an opening bracket in an expression. {0}: bracket, {1}: expression</summary>
        UnbalancedBracket,

        /// <summary>An unterminated string literal in an expression. {0}: expression</summary>
        UnterminatedStringLiteral,

        /// <summary>An opening bracket that is not closed in an expression. {0}: bracket, {1}: expression</summary>
        UnclosedBracket,

        /// <summary>Closing braces without opening braces. {0}: braces</summary>
        UnexpectedClosingBraces,

        /// <summary>A malformed tag that stays in the document as text. {0}: tag</summary>
        InvalidTag,

        /// <summary>An opening tag that is never closed with <c>}}</c>. {0}: tag</summary>
        TagNotTerminated,

        /// <summary>A switch block tag that could not be parsed. {0}: tag content</summary>
        InvalidSwitchBlockSyntax,

        /// <summary>A case block tag that could not be parsed. {0}: tag content</summary>
        InvalidCaseBlockSyntax,

        /// <summary>A case or default block that is not nested in a switch block. {0}: case expression or "default"</summary>
        CaseNotNestedInSwitch,

        /// <summary>An unknown inline keyword. {0}: tag</summary>
        InvalidInlineKeyword,

        /// <summary>A dynamic table block does not contain a table with a header and a data row.</summary>
        DynamicTableStructureInvalid,

        /// <summary>A dynamic table block does not contain exactly one child block.</summary>
        DynamicTableSingleChild
    }
}
