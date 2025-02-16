namespace DocxTemplater
{
    internal enum PatternType
    {
        None,
        Condition,
        ConditionEnd,
        CollectionStart,
        CollectionSeparator,
        CollectionEnd,
        Variable,
        ConditionElse,
        InlineKeyWord
    }
}
