namespace DocxTemplater
{
    internal enum PatternType
    {
        None,
        Condition,
        ConditionEnd,
        CollectionStart,
        DynamicTable,
        CollectionSeparator,
        CollectionEnd,
        Variable,
        ConditionElse,
        InlineKeyWord
    }
}
