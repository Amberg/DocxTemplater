using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace DocxTemplater
{
    internal record PatternMatch(Match Match, PatternType Type, string Condition, string Prefix, string Variable,
        string Formatter, string[] Arguments, int Index, int Length);

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

        private static readonly Regex PatternRegex = new(@"\{\{
                                                                (?:   
                                                                    (?<else>else) |
                                                                    (?:
                                                                        (?<prefix>[\/\#])? #prefix
                                                                        (?:
                                                                            (?<varname>[a-zA-Z0-9\._]+) #variable name
                                                                            | #or
                                                                            (?<condition>[a-zA-Z0-9+\-*\/><=\s\.]+) #condition
                                                                        )?
                                                                    )
                                                                )
                                                            \}
                                                            (?::
                                                                (?<formatter>[a-zA-z0-9]+) #formatter
                                                                (?:\( #arguments with brackets
                                                                    (?:
                                                                        (?:,?
                                                                            (?:   (?<arg>[a-zA-Z0-9\s-\\/_-]+) | (?:'(?<arg>[a-zA-Z0-9\s-\\/_-]*)')   )
                                                                        )* 
                                                                    )
                                                                \))?
                                                            )?
                                                            \}
                                                            ", RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace);
        public static IEnumerable<PatternMatch> FindSyntaxPatterns(string text)
        {
            var matches = PatternRegex.Matches(text);
            var result = new List<PatternMatch>(matches.Count);
            foreach (Match match in matches)
            {
                if (match.Groups["else"].Success)
                {
                    result.Add(new PatternMatch(match, PatternType.ConditionElse, null, null, null, null, null, match.Index, match.Length));
                }
                else

                if (match.Groups["condition"].Success)
                {
                    result.Add(new PatternMatch(match, PatternType.Condition, match.Groups["condition"].Value, null, null, null, null, match.Index, match.Length));
                }
                else if (match.Groups["prefix"].Success)
                {
                    if (match.Groups["prefix"].Value == "#")
                    {
                        result.Add(new PatternMatch(match, PatternType.CollectionStart, null, match.Groups["prefix"].Value, match.Groups["varname"].Value, match.Groups["formatter"].Value, match.Groups["arg"].Value.Split(','), match.Index, match.Length));
                    }
                    else if (!match.Groups["varname"].Success)
                    {
                        result.Add(new PatternMatch(match, PatternType.ConditionEnd, null, null, null, null, null, match.Index, match.Length));
                    }
                    else
                    {
                        result.Add(new PatternMatch(match, PatternType.CollectionEnd, null, match.Groups["prefix"].Value, match.Groups["varname"].Value, match.Groups["formatter"].Value, match.Groups["arg"].Value.Split(','), match.Index, match.Length));
                    }
                }
                else
                {
                    var argGroup = match.Groups["arg"];
                    var arguments = argGroup.Success ? argGroup.Captures.Select(x => x.Value).ToArray() : Array.Empty<string>();
                    result.Add(new PatternMatch(match, PatternType.Variable, null, null, match.Groups["varname"].Value, match.Groups["formatter"].Value, arguments, match.Index, match.Length));
                }
            }
            return result;
        }
    }
}
