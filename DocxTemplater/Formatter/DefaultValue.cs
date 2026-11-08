using System;

namespace DocxTemplater.Formatter
{
    /// <summary>
    /// <c>{{value}:default('text')}</c> / <c>{{value}:ifempty('text')}</c>: the text if the value is <c>null</c>, empty or
    /// whitespace, otherwise the value formatted as without a formatter. The text is literal - it is not evaluated as a
    /// template or an expression. Handled before the value reaches a formatter, since formatters never see <c>null</c>.
    /// </summary>
    internal static class DefaultValue
    {
        public static bool IsDefaultFormatter(string formatter)
        {
            return formatter != null
                   && (formatter.Equals("default", StringComparison.OrdinalIgnoreCase) || formatter.Equals("ifempty", StringComparison.OrdinalIgnoreCase));
        }

        public static bool IsEmpty(object value)
        {
            return value == null || (value is string text && string.IsNullOrWhiteSpace(text));
        }

        /// <summary>
        /// The replacement text. Unquoted text is split at the commas like any formatter argument, so the parts are
        /// joined with <c>", "</c>: <c>default(a,b)</c> is <c>a, b</c>. Quoted text is kept as written.
        /// </summary>
        public static string Text(string[] arguments)
        {
            return arguments == null ? string.Empty : string.Join(", ", arguments);
        }

        /// <summary>
        /// If <paramref name="match"/> uses the default formatter: <c>true</c> and the replacement text when
        /// <paramref name="value"/> is empty; otherwise <c>false</c> and the match without the formatter, so the value is
        /// formatted as without it (including a default formatter of the model).
        /// </summary>
        public static bool TryGetReplacement(ref PatternMatch match, object value, out string replacement)
        {
            replacement = null;
            if (!IsDefaultFormatter(match.Formatter))
            {
                return false;
            }
            if (IsEmpty(value))
            {
                replacement = Text(match.Arguments);
                return true;
            }
            match = match with { Formatter = null, Arguments = Array.Empty<string>() };
            return false;
        }
    }
}
