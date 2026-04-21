using System.Collections.Generic;

namespace DocxTemplater
{
    internal static class HelperFunctions
    {

        public static string SanitizeQuotes(string input)
        {
            return input.Trim()
                .Replace('\'', '"')
                .Replace('\u2018', '"') // Left Single
                .Replace('\u2019', '"') // Right Single
                .Replace('\u201C', '"') // Left Double
                .Replace('\u201D', '"') // Right Double
                .Replace('\u00AB', '"') // Left-Pointing
                .Replace('\u00BB', '"');// Right-Pointing
        }

        /// <summary>
        /// Parses arguments in form foo:bar,foo2:bar2
        /// </summary>
        /// <param name="arguments"></param>
        public static Dictionary<string, string> ParseArguments(string[] arguments)
        {
            var result = new Dictionary<string, string>(arguments.Length);
            foreach (var arg in arguments)
            {
                var parts = arg.Split(':');
                if (parts.Length != 2)
                {
                    throw new OpenXmlTemplateException("Arguments must be in the form foo:bar");
                }
                var val = SanitizeQuotes(parts[1]).Replace("\"", "").Trim();
                result[parts[0]] = val;
            }
            return result;
        }
    }
}