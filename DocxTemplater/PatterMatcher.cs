using System;
using System.Collections.Generic;
using System.Linq;
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

    internal static class PatternMatcher
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

        private static readonly Regex PatternRegex = new(@"\{\s*(?<condMarker>\?\s*)?\{\s*
        (?:   
            (?<separator>:\s*s\s*:) |
            (?<else>(?:else|:(?!\s*[^\s}]+))) |  # Match : only if nothing but whitespace follows before }
            (?(condMarker)
                (?<condition>[^}]+) # allow any character except }
                |
                (?:
                    (?<prefix>[\/\#:@]{1,2})?
                    (?:
                        (?<varname>  # Constrain Switch/Case capturing to strictly known prefix variables
                            (?i:switch|s|case|c)
                            (?:
                                \s*:\s* # match colon for switch/case
                                (?:
                                    (?:'[^']*') | # match single quotes
                                    (?:""[^""]*"") | # match double quotes
                                    [\p{L}\p{N}\._()]+ # match unquoted values
                                )
                            )?
                        )
                        |
                        (?<varname> # Range loop pattern: identifier:count
                            [\p{L}\p{N}\._]+
                            :
                            [\p{L}\p{N}\._]+
                        )
                        |
                        (?<varname> # the default variable matcher
                            [\p{L}\p{N}\._]+ # match variable name
                        )
                    )?
                ) 
            )
        )
        \s*\}
        (?::
            (?<formatter>[\p{L}\p{N}]+)
            (?:\(
                (?:
                   (?:
                       (?:'(?<arg>(?:(?:\\['])|[^'}])*?)'|""""(?<arg>(?:(?:\\[""""])|[^""""}])*?)"""") # Allow any character except ' and } in quoted strings
                        |
                        (?<arg>[^,)}]+) # Allow any character except delimiters in unquoted strings
                    )(?:\s*,\s*)?
                )*
                \))?
        )?
        \s*\}
        ", RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace, TimeSpan.FromSeconds(1));

        public static IEnumerable<PatternMatch> FindSyntaxPatterns(string text)
        {
            try
            {
                if (text == null)
                {
                    return Array.Empty<PatternMatch>();
                }

                var matches = PatternRegex.Matches(text);
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
                            var patternType = PatternType.CollectionEnd;
                            var trimmedVarname = varname.Trim();
                            if (trimmedVarname.Equals("switch", StringComparison.OrdinalIgnoreCase) || trimmedVarname.Equals("s", StringComparison.OrdinalIgnoreCase))
                            {
                                patternType = PatternType.SwitchEnd;
                            }
                            else if (trimmedVarname.Equals("case", StringComparison.OrdinalIgnoreCase) || trimmedVarname.Equals("c", StringComparison.OrdinalIgnoreCase))
                            {
                                patternType = PatternType.CaseEnd;
                            }
                            else if (trimmedVarname.Equals("default", StringComparison.OrdinalIgnoreCase) || trimmedVarname.Equals("d", StringComparison.OrdinalIgnoreCase))
                            {
                                patternType = PatternType.DefaultEnd;
                            }

                            result.Add(new PatternMatch(match, patternType, null,
                                match.Groups["prefix"].Value, match.Groups["varname"].Value,
                                match.Groups["formatter"].Value, match.Groups["arg"].Value.Split(','), match.Index,
                                match.Length));
                        }
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
    }
}
