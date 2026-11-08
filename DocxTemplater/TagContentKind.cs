namespace DocxTemplater
{
    /// <summary>The part of a tag passed to <see cref="ITagContentPreprocessor.Process"/>.</summary>
    public enum TagContentKind
    {
        /// <summary>
        /// A condition (<c>{?{a &gt; 1}}</c> → <c>a &gt; 1</c>), an expression including its parentheses
        /// (<c>{{(a + 1)}}</c> → <c>(a + 1)</c>) or the expression of a switch or case (<c>{{#case:'A'}}</c> → <c>'A'</c>).
        /// </summary>
        Expression,

        /// <summary>
        /// The arguments of a formatter as written between the parentheses, including quotes and commas
        /// (<c>{{x}:default('Smith, Jones')}</c> → <c>'Smith, Jones'</c>). The result is split into arguments by the
        /// usual rules, so a preprocessor can also replace the quotes that delimit an argument.
        /// </summary>
        FormatterArguments
    }
}
