using System.Text;

namespace DocxTemplater.Text
{
    /// <summary>
    /// The encoding of <see cref="TextEscaping.Html"/>: only what HTML requires - <c>&amp; &lt; &gt; &quot; &#39;</c> -
    /// in text and in quoted attribute values. Umlauts and accented letters stay as they are (the output is Unicode).
    /// </summary>
    internal static class HtmlEncoding
    {
        private static readonly char[] SpecialCharacters = ['&', '<', '>', '"', '\''];

        public static string Encode(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOfAny(SpecialCharacters) < 0)
            {
                return text ?? string.Empty;
            }
            var result = new StringBuilder(text.Length + 16);
            foreach (var c in text)
            {
                switch (c)
                {
                    case '&':
                        result.Append("&amp;");
                        break;
                    case '<':
                        result.Append("&lt;");
                        break;
                    case '>':
                        result.Append("&gt;");
                        break;
                    case '"':
                        result.Append("&quot;");
                        break;
                    case '\'':
                        result.Append("&#39;");
                        break;
                    default:
                        result.Append(c);
                        break;
                }
            }
            return result.ToString();
        }
    }
}
