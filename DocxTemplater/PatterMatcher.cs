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
                        |
                        (?<expression> # the expression matcher
                            \(
                            (?>[^()]+|(?<o>\()|(?<-o>\)))*
                            \)
                            (?(o)(?!))
                        )
                    )?
                ) 
            )
        )
        \s*\}
        (?::
            (?<formatter>[\p{L}\p{N}]+)
            (?:\(
                (?<argList>(?:
                   (?:
                       (?:'(?<arg>(?:(?:\\['])|[^'}])*?)'|""""(?<arg>(?:(?:\\[""""])|[^""""}])*?)"""") # Allow any character except ' and } in quoted strings
                        |
                        (?<arg>[^,)}]+) # Allow any character except delimiters in unquoted strings
                    )(?:\s*,\s*)?
                )*)
                \))?
        )?
        \s*\}
        ", RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace, TimeSpan.FromSeconds(1));

        private static readonly Regex ArgumentRegex = new(@"\G(?:'(?<arg>(?:(?:\\['])|[^'])*?)'|""""(?<arg>(?:(?:\\[""""])|[^""""])*?)""""|(?<arg>[^,]+))(?:\s*,\s*)?",
            RegexOptions.Compiled, TimeSpan.FromSeconds(1));

        public static IEnumerable<PatternMatch> FindSyntaxPatterns(string text)
        {
            return FindSyntaxPatterns(text, null, null);
        }

        /// <summary>
        /// Finds all syntax patterns in <paramref name="text"/>; an invalid pattern throws an
        /// <see cref="OpenXmlTemplateException"/> in the language of <paramref name="settings"/>.
        /// </summary>
        public static IEnumerable<PatternMatch> FindSyntaxPatterns(string text, ProcessSettings settings)
        {
            return FindSyntaxPatterns(text, null, settings);
        }

        /// <summary>
        /// Finds all syntax patterns in <paramref name="text"/>. If <paramref name="onInvalidPattern"/> is null,
        /// an invalid pattern throws an <see cref="OpenXmlTemplateException"/> in the language of
        /// <paramref name="settings"/> (<c>null</c> for English); otherwise the error code and its arguments are
        /// reported to the callback and the pattern is skipped, so all errors of a text can be collected.
        /// </summary>
        public static IEnumerable<PatternMatch> FindSyntaxPatterns(string text, Action<Match, TemplateErrorCode, object[]> onInvalidPattern, ProcessSettings settings)
        {
            return FindSyntaxPatterns(text, onInvalidPattern, settings, TagContentPreprocessors.None);
        }

        /// <summary>The content of each tag is passed through <paramref name="preprocessors"/>, its position is not changed.</summary>
        /// <inheritdoc cref="FindSyntaxPatterns(string, Action{Match, TemplateErrorCode, object[]}, ProcessSettings)"/>
        public static IEnumerable<PatternMatch> FindSyntaxPatterns(string text, Action<Match, TemplateErrorCode, object[]> onInvalidPattern, ProcessSettings settings, TagContentPreprocessors preprocessors)
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
                                ReportInvalid(onInvalidPattern, settings, match, TemplateErrorCode.InvalidSyntax, match.Value);
                                continue;
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
                            if (varname == null || !varname.Equals("ignore", StringComparison.CurrentCultureIgnoreCase))
                            {
                                ReportInvalid(onInvalidPattern, settings, match, TemplateErrorCode.InvalidSyntax, match.Value);
                                continue;
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
                                ReportInvalid(onInvalidPattern, settings, match, TemplateErrorCode.RangeLoopVariableNameRequired, match.Value);
                                continue;
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
                                ReportInvalid(onInvalidPattern, settings, match, TemplateErrorCode.UseGenericClosingTag, match.Value);
                                continue;
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
                            ? argGroup.Captures.Cast<Capture>().Select(x => x.Value?.Replace("\\'", "'")).ToArray()
                            : Array.Empty<string>();

                        // we want to extract the expression string but keep the original match structure similar to Variable
                        // the expression group includes the outer parenthesis `(a > 5)`
                        // let's pass the whole captured expression as the variable name for now, the block can trim if needed.
                        result.Add(new PatternMatch(match, PatternType.Expression, null, null,
                            match.Groups["expression"].Value,
                            match.Groups["formatter"].Value, arguments, match.Index, match.Length));
                    }
                    else if (match.Groups["varname"].Success)
                    {
                        var argGroup = match.Groups["arg"];
                        var arguments = argGroup.Success
                            ? argGroup.Captures.Cast<Capture>().Select(x => x.Value?.Replace("\\'", "'")).ToArray()
                            : Array.Empty<string>();
                        result.Add(new PatternMatch(match, PatternType.Variable, null, null,
                            match.Groups["varname"].Value,
                            match.Groups["formatter"].Value, arguments, match.Index, match.Length));
                    }
                    else
                    {
                        ReportInvalid(onInvalidPattern, settings, match, TemplateErrorCode.InvalidSyntax, match.Value);
                    }
                }

                if (preprocessors != null && !preprocessors.IsEmpty)
                {
                    for (int i = 0; i < result.Count; i++)
                    {
                        result[i] = Preprocess(result[i], preprocessors);
                    }
                }
                return result;
            }
            catch (RegexMatchTimeoutException)
            {
                throw OpenXmlTemplateException.Create(settings, TemplateErrorCode.PatternMatchTimeout, text);
            }
        }

        private static PatternMatch Preprocess(PatternMatch match, TagContentPreprocessors preprocessors)
        {
            switch (match.Type)
            {
                case PatternType.Condition:
                    match = match with { Condition = preprocessors.Process(match.Condition, TagContentKind.Expression) };
                    break;
                case PatternType.Expression:
                    match = match with { Variable = preprocessors.Process(match.Variable, TagContentKind.Expression) };
                    break;
                case PatternType.Switch:
                case PatternType.Case:
                    var colon = match.Variable.IndexOf(':');
                    if (colon >= 0)
                    {
                        match = match with { Variable = match.Variable[..(colon + 1)] + preprocessors.Process(match.Variable[(colon + 1)..], TagContentKind.Expression) };
                    }
                    break;
            }

            var argList = match.Match.Groups["argList"];
            if (match.Type is PatternType.Variable or PatternType.Expression && argList.Success)
            {
                var processed = preprocessors.Process(argList.Value, TagContentKind.FormatterArguments);
                if (processed != argList.Value)
                {
                    match = match with { Arguments = SplitArguments(processed) };
                }
            }
            return match;
        }

        /// <summary>Splits a preprocessed argument list by the rules of the argument part of <see cref="PatternRegex"/>.</summary>
        private static string[] SplitArguments(string argList)
        {
            var arguments = new List<string>();
            var m = ArgumentRegex.Match(argList);
            while (m.Success && m.Length > 0)
            {
                arguments.Add(m.Groups["arg"].Value.Replace("\\'", "'"));
                m = m.NextMatch();
            }
            return [.. arguments];
        }

        private static void ReportInvalid(Action<Match, TemplateErrorCode, object[]> onInvalidPattern, ProcessSettings settings, Match match, TemplateErrorCode code, params object[] args)
        {
            if (onInvalidPattern == null)
            {
                throw OpenXmlTemplateException.Create(settings, code, args);
            }
            onInvalidPattern(match, code, args);
        }
    }
}
