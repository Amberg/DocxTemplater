using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace DocxTemplater
{
    internal record PatternMatch(
        Match Match,
        PatternType Type,
        string Condition,
        string Prefix,
        string Variable,
        string Formatter,
        string[] Arguments,
        int Index,
        int Length);

    internal class PatternMatcher
    {
        /* matches:
        {{a > 5}} -- conditional - only capture
        {{/}} -- conditional end
        {{#images}} -- loop start
        {{/images}} -- loop end
        {{images}} -- variable
        {{images}:foo()} -- variable with formatter
        {{images}:foo(arg1,arg2)} -- variable with formatter and arguments
         */

        private readonly Regex _regex;
        private readonly string _openDelimiter;
        private readonly string _closeDelimiter;

        internal static PatternMatcher Default { get; } = new("{{", "}}");

        internal PatternMatcher(string openDelimiter, string closeDelimiter)
        {
            ValidateDelimiters(openDelimiter, closeDelimiter);
            _openDelimiter = openDelimiter;
            _closeDelimiter = closeDelimiter;
            _regex = BuildRegex(openDelimiter, closeDelimiter);
        }

        /// <summary>
        /// Builds the synthetic string used internally to parse a formatter spec from metadata,
        /// e.g. "{{x}:F(n)}" for the default delimiters, or "&lt;&lt;x&gt;:F(n)&gt;" for "&lt;&lt;"/"&gt;&gt;" delimiters.
        /// </summary>
        internal string BuildFormatterCheckSyntax(string formatterSpec)
        {
            return _openDelimiter + "x" + _closeDelimiter[..^1] + ":" + formatterSpec + _closeDelimiter[^1..];
        }

        internal IEnumerable<PatternMatch> FindSyntaxPatterns(string text)
        {
            try
            {
                if (text == null)
                {
                    return Array.Empty<PatternMatch>();
                }

                var matches = _regex.Matches(text);
                var result = new List<PatternMatch>(matches.Count);
                foreach (Match match in matches)
                {
                    if (match.Groups["separator"].Success)
                    {
                        result.Add(new PatternMatch(match, PatternType.CollectionSeparator, null, null, null, null,
                            null,
                            match.Index, match.Length));
                    }
                    else if (match.Groups["else"].Success)
                    {
                        result.Add(new PatternMatch(match, PatternType.ConditionElse, null, null, null, null, null,
                            match.Index, match.Length));
                    }
                    else if (match.Groups["condition"].Success)
                    {
                        result.Add(new PatternMatch(match, PatternType.Condition, match.Groups["condition"].Value, null,
                            null, null, null, match.Index, match.Length));
                    }
                    else if (match.Groups["prefix"].Success)
                    {
                        string varname = match.Groups["varname"].Success ? match.Groups["varname"].Value : null;
                        var prefix = match.Groups["prefix"].Value;
                        if (prefix == ":")
                        {
                            if (varname == null)
                            {
                                throw new OpenXmlTemplateException($"Invalid syntax '{match.Value}'");
                            }

                            if (varname.Equals("ignore", StringComparison.CurrentCultureIgnoreCase))
                            {
                                result.Add(new PatternMatch(match, PatternType.IgnoreBlock, null, prefix, varname, null,
                                    null, match.Index, match.Length));
                            }
                            else
                            {
                                result.Add(new PatternMatch(match, PatternType.InlineKeyWord, null, prefix, varname,
                                    null,
                                    null, match.Index, match.Length));
                            }
                        }
                        else if (prefix == "/:")
                        {
                            if (!varname.Equals("ignore", StringComparison.CurrentCultureIgnoreCase))
                            {
                                throw new OpenXmlTemplateException($"Invalid syntax '{match.Value}'");
                            }
                            result.Add(new PatternMatch(match, PatternType.IgnoreEnd, null, prefix, varname, null, null, match.Index, match.Length));

                        }
                        else if (prefix == "#")
                        {
                            var patternType = PatternType.CollectionStart;
                            if (varname != null)
                            {
                                var trimmedVarname = varname.Trim();
                                if (trimmedVarname.StartsWith("switch:", StringComparison.OrdinalIgnoreCase) || trimmedVarname.StartsWith("s:", StringComparison.OrdinalIgnoreCase))
                                {
                                    patternType = PatternType.Switch;
                                }
                                else if (trimmedVarname.StartsWith("case:", StringComparison.OrdinalIgnoreCase) || trimmedVarname.StartsWith("c:", StringComparison.OrdinalIgnoreCase))
                                {
                                    patternType = PatternType.Case;
                                }
                                else if (trimmedVarname.Equals("default", StringComparison.OrdinalIgnoreCase) || trimmedVarname.Equals("d", StringComparison.OrdinalIgnoreCase))
                                {
                                    patternType = PatternType.Default;
                                }
                                else if (trimmedVarname.Contains(':'))
                                {
                                    // Colon syntax is only valid for switch/case keywords, not for regular collections
                                    continue;
                                }
                            }

                            result.Add(new PatternMatch(match, patternType, null,
                                match.Groups["prefix"].Value, match.Groups["varname"].Value,
                                match.Groups["formatter"].Value, match.Groups["arg"].Value.Split(','), match.Index,
                                match.Length));
                        }
                        else if (prefix == "@")
                        {
                            if (string.IsNullOrWhiteSpace(varname))
                            {
                                throw new OpenXmlTemplateException($"Invalid range loop syntax '{match.Value}' - variable name is required");
                            }

                            result.Add(new PatternMatch(match, PatternType.RangeStart, null,
                                match.Groups["prefix"].Value, match.Groups["varname"].Value,
                                match.Groups["formatter"].Value, match.Groups["arg"].Value.Split(','), match.Index,
                                match.Length));
                        }
                        else if (varname == null)
                        {
                            result.Add(new PatternMatch(match, PatternType.ConditionEnd, null, null, null, null, null,
                                match.Index, match.Length));
                        }
                        else
                        {
                            var trimmedVarname = varname.Trim();
                            if (trimmedVarname.Equals("switch", StringComparison.OrdinalIgnoreCase) ||
                                trimmedVarname.Equals("s", StringComparison.OrdinalIgnoreCase) ||
                                trimmedVarname.Equals("case", StringComparison.OrdinalIgnoreCase) ||
                                trimmedVarname.Equals("c", StringComparison.OrdinalIgnoreCase) ||
                                trimmedVarname.Equals("default", StringComparison.OrdinalIgnoreCase) ||
                                trimmedVarname.Equals("d", StringComparison.OrdinalIgnoreCase))
                            {
                                throw new OpenXmlTemplateException($"Invalid syntax '{match.Value}'. Use '{{{{/}}}}' instead.");
                            }
                            var patternType = PatternType.CollectionEnd;
                            result.Add(new PatternMatch(match, patternType, null,
                                match.Groups["prefix"].Value, match.Groups["varname"].Value,
                                match.Groups["formatter"].Value, match.Groups["arg"].Value.Split(','), match.Index,
                                match.Length));
                        }
                    }
                    else if (match.Groups["expression"].Success)
                    {
                        var argGroup = match.Groups["arg"];
                        var arguments = argGroup.Success
                            ? argGroup.Captures.Select(x => x.Value?.Replace("\\'", "'")).ToArray()
                            : Array.Empty<string>();

                        result.Add(new PatternMatch(match, PatternType.Expression, null, null,
                            match.Groups["expression"].Value,
                            match.Groups["formatter"].Value, arguments, match.Index, match.Length));
                    }
                    else if (match.Groups["varname"].Success)
                    {
                        var argGroup = match.Groups["arg"];
                        var arguments = argGroup.Success
                            ? argGroup.Captures.Select(x => x.Value?.Replace("\\'", "'")).ToArray()
                            : Array.Empty<string>();
                        result.Add(new PatternMatch(match, PatternType.Variable, null, null,
                            match.Groups["varname"].Value,
                            match.Groups["formatter"].Value, arguments, match.Index, match.Length));
                    }
                    else
                    {
                        throw new OpenXmlTemplateException($"Invalid syntax '{match.Value}'");
                    }
                }

                return result;
            }
            catch (RegexMatchTimeoutException)
            {
                throw new OpenXmlTemplateException($"Invalid syntax '{text}' - match timeout");
            }
        }

        private static Regex BuildRegex(string open, string close)
        {
            // Split delimiter into outer char and inner chars:
            //   open[0]     = outer open  (e.g. '{' for "{{")
            //   open[1..]   = inner open  (e.g. '{' for "{{")
            //   close[..^1] = inner close (e.g. '}' for "}}")
            //   close[^1..] = outer close (e.g. '}' for "}}")
            //
            // For "<<" / ">>": outer='<', inner='<', inner-close='>', outer-close='>'
            // Conditional syntax: outerOpen + '?' + innerOpen  (e.g. '{?{' or '<?<')
            string eo = Regex.Escape(open[0].ToString());
            string ei = Regex.Escape(open[1..]);
            string c0 = Regex.Escape(close[..^1]);
            string c1 = Regex.Escape(close[^1..]);
            string cc0 = EscapeForCharClass(close[..^1]);
            string dq = "\"\"";

            // Build the full regex pattern. RegexOptions.IgnorePatternWhitespace is used so
            // whitespace outside character classes is ignored, allowing readability via newlines.
            string pattern =
                eo + @"\s*(?<condMarker>\?\s*)?" + ei + @"\s*" +
                @"(?:" +
                    @"(?<separator>:\s*s\s*:) |" +
                    @"(?<else>(?:else|:(?!\s*[^\s" + cc0 + @"]+))) |" +
                    @"(?(condMarker)" +
                        @"(?<condition>[^" + cc0 + @"]+)" +
                        @"|" +
                        @"(?:" +
                            @"(?<prefix>[\/\#:@]{1,2})?" +
                            @"(?:" +
                                @"(?<varname>" +
                                    @"(?i:switch|s|case|c)" +
                                    @"(?:" +
                                        @"\s*:\s*" +
                                        @"(?:" +
                                            @"(?:'[^']*') |" +
                                            @"(?:""[^""]*"") |" +
                                            @"[\p{L}\p{N}\._()]+" +
                                        @")" +
                                    @")?" +
                                @")" +
                                @"|" +
                                @"(?<varname>" +
                                    @"[\p{L}\p{N}\._]+" +
                                    @":" +
                                    @"[\p{L}\p{N}\._]+" +
                                @")" +
                                @"|" +
                                @"(?<varname>" +
                                    @"[\p{L}\p{N}\._]+" +
                                @")" +
                                @"|" +
                                @"(?<expression>" +
                                    @"\(" +
                                    @"(?>[^()]+|(?<o>\()|(?<-o>\)))*" +
                                    @"\)" +
                                    @"(?(o)(?!))" +
                                @")" +
                            @")?" +
                        @")" +
                    @")" +
                @")" +
                @"\s*" + c0 +
                @"(?::" +
                    @"(?<formatter>[\p{L}\p{N}]+)" +
                    @"(?:\(" +
                        @"(?:" +
                            @"(?:" +
                                "(?:'(?<arg>(?:(?:\\\\['])|[^'" + cc0 + "])*?)'|" +
                                dq + "(?<arg>(?:(?:\\\\[" + dq + "])|[^" + dq[0] + cc0 + "])*?)" + dq +
                                @")" +
                                @"|" +
                                "(?<arg>[^,)" + cc0 + "]+)" +
                            @")(?:\s*,\s*)?" +
                        @")*" +
                        @"\))?" +
                @")?" +
                @"\s*" + c1;

            return new Regex(pattern, RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace, TimeSpan.FromSeconds(1));
        }

        private static void ValidateDelimiters(string open, string close)
        {
            if (string.IsNullOrEmpty(open))
            {
                throw new ArgumentException("OpenDelimiter must not be null or empty.", nameof(open));
            }
            if (open.Length < 2)
            {
                throw new ArgumentException("OpenDelimiter must be at least 2 characters.", nameof(open));
            }
            if (string.IsNullOrEmpty(close))
            {
                throw new ArgumentException("CloseDelimiter must not be null or empty.", nameof(close));
            }
            if (close.Length < 2)
            {
                throw new ArgumentException("CloseDelimiter must be at least 2 characters.", nameof(close));
            }
            if (open == close)
            {
                throw new ArgumentException("OpenDelimiter and CloseDelimiter must not be the same.");
            }
        }

        private static string EscapeForCharClass(string s)
        {
            var sb = new StringBuilder(s.Length * 2);
            foreach (char c in s)
            {
                if (c is '\\' or ']' or '^' or '-')
                {
                    sb.Append('\\');
                }
                sb.Append(c);
            }
            return sb.ToString();
        }
    }
}
