using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using DocxTemplater.Blocks;
using DynamicExpresso;

namespace DocxTemplater
{
    /// <summary>
    /// The parsed conditions and expressions of all templates created with the same settings. A script is pre-processed
    /// once (implicit scope <c>.Name</c>, <c>_Idx</c> / <c>_Length</c>, model identifiers) and parsed once per
    /// combination of value types; the model values are parameters of the compiled delegate, never part of it.
    /// Thread-safe.
    /// </summary>
    internal sealed class ScriptCache
    {
        // shared by all threads: "only get properties, Parse and Eval methods are thread safe", as long as the interpreter
        // is not configured further (DynamicExpresso README, "Performance and multithreading"); DetectIdentifiers only reads it
        private static readonly Interpreter SharedInterpreter = new Interpreter().EnableAssignment(AssignmentOperators.None);

        private static readonly Regex RegexWordStartingWithDot = new(@"
                                                                    (?x)                         # Enable verbose mode - allows comments and whitespace in pattern
                                                                    (?:                          # Non-capturing group for all possible prefixes:
                                                                        ^                        # Either start of string
                                                                        |
                                                                        (?<unary>[+\-!])         # Capture unary operators (+, -, !) in 'unary' group
                                                                        |
                                                                        (?<=[^.\p{L}\p{N}_?\]])  # Position after any char that's not a dot, letter, number, underscore, question mark (for null-conditional)
                                                                                                 # or closing bracket (a dot after ']' is member access on an index result, not an implicit scope)
                                                                                                 # (lookbehind - ensures we don't break existing identifiers or ?.)

                                                                        (?<!\(.*\))              # Also try to start after any function call with zero to many arguments.
                                                                                                 # NOTE: While this isn't industrial grade combinatorial parsing, the results seem fine.
                                                                    )
                                                                    (?<dots>\.+)                 # Capture one or more dots in 'dots' group
                                                                    (?<prop>                     # Start 'prop' group for the property name
                                                                        [\p{L}\p{N}_]*           # Any number of letters, numbers, or underscores
                                                                    )
                                                                    ",
            RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace, TimeSpan.FromMilliseconds(500)
        );

        private static readonly Regex RegexListKeywords = new(@"(?<dots>\.*)(?<variable>\p{L}[\p{L}\p{N}_]*(?:\.\p{L}[\p{L}\p{N}_]*)*)?\.(?<keyword>_Idx|_Length)"
            , RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace, TimeSpan.FromMilliseconds(500));

        private readonly ConcurrentDictionary<string, PreparedScript> m_scripts = new(StringComparer.Ordinal);

        internal enum BindingKind
        {
            /// <summary>A model or scope variable referenced by name, e.g. <c>ds</c> in <c>ds.Count &gt; 1</c>.</summary>
            Identifier,

            /// <summary>The item of an enclosing loop, <c>.Name</c> / <c>..Name</c>.</summary>
            Scope,

            /// <summary>A loop keyword resolved by its path, e.g. <c>ds.Items._Idx</c>.</summary>
            Path
        }

        /// <summary>A parameter of a script and where its value comes from.</summary>
        internal sealed record Binding(string Name, BindingKind Kind, string Path, int DotCount);

        /// <exception cref="RegexMatchTimeoutException">The script could not be pre-processed.</exception>
        public PreparedScript Prepare(string script)
        {
            if (m_scripts.TryGetValue(script, out var prepared))
            {
                return prepared;
            }
            return m_scripts.GetOrAdd(script, CreatePreparedScript(script));
        }

        private static PreparedScript CreatePreparedScript(string script)
        {
            var bindings = new List<Binding>();
            var text = HelperFunctions.SanitizeQuotes(script);
            text = RegexWordStartingWithDot.Replace(text, m => OnVariableReplace(m, bindings));
            text = RegexListKeywords.Replace(text, m => OnSpecialKeyWordsReplace(m, bindings));

            var identifiers = SharedInterpreter.DetectIdentifiers(text);
            foreach (var identifier in identifiers.UnknownIdentifiers)
            {
                if (bindings.All(x => x.Name != identifier))
                {
                    bindings.Add(new Binding(identifier, BindingKind.Identifier, identifier, 0));
                }
            }
            return new PreparedScript(text, bindings);
        }

        private static string OnVariableReplace(Match match, List<Binding> bindings)
        {
            string unary = match.Groups["unary"].Value; // Preserve unary operators if present
            string prop = match.Groups["prop"].Value; // Preserve the word after dots
            var dotCount = match.Groups["dots"].Value.Length;

            if (prop == LoopBlock.LoopIndexVariable || prop == LoopBlock.LoopLengthVariable)
            {
                var varName = $"__s{dotCount}_{prop}";
                AddBinding(bindings, new Binding(varName, BindingKind.Path, match.Value, dotCount));
                return varName;
            }
            else
            {
                var varName = $"__s{dotCount}_"; // choose a variable name that is unlikely to be used by the user
                AddBinding(bindings, new Binding(varName, BindingKind.Scope, null, dotCount));
                if (!string.IsNullOrWhiteSpace(prop))
                {
                    varName += $".{prop}";
                }

                return $"{unary}{varName}";
            }
        }

        private static string OnSpecialKeyWordsReplace(Match match, List<Binding> bindings)
        {
            var dots = match.Groups["dots"].Value.Length;
            var keyWord = match.Groups["keyword"];
            var varName = $"__s{dots}_{keyWord}";
            AddBinding(bindings, new Binding(varName, BindingKind.Path, match.Value, dots));
            return varName;
        }

        // a name used twice gets the value of its last use, as a variable set twice on the interpreter would
        private static void AddBinding(List<Binding> bindings, Binding binding)
        {
            var index = bindings.FindIndex(x => x.Name == binding.Name);
            if (index >= 0)
            {
                bindings[index] = binding;
            }
            else
            {
                bindings.Add(binding);
            }
        }

        /// <summary>A pre-processed script with the delegates parsed for it so far, by return type and parameter types.</summary>
        internal sealed class PreparedScript
        {
            private readonly ConcurrentDictionary<SignatureKey, Func<object[], object>> m_functions = new();
            // only the message: one exception instance thrown on several threads would share its stack trace;
            // per UI culture, the language of the parser's message
            private readonly ConcurrentDictionary<(SignatureKey, CultureInfo), string> m_parseErrors = new();

            public PreparedScript(string text, IReadOnlyList<Binding> bindings)
            {
                Text = text;
                Bindings = bindings;
            }

            /// <summary>The script as it is parsed, with the implicit scope and loop keywords replaced by parameters.</summary>
            public string Text { get; }

            /// <summary>The parameters of the script, in the order of the values passed to the delegate.</summary>
            public IReadOnlyList<Binding> Bindings { get; }

            /// <returns><c>false</c> and the message of the parser if the script cannot be parsed with these types.</returns>
            public bool TryGetFunction(Type returnType, Type[] parameterTypes, out Func<object[], object> function, out string parseError)
            {
                var key = new SignatureKey(returnType, parameterTypes);
                var errorKey = (key, CultureInfo.CurrentUICulture);
                parseError = null;
                if (m_functions.TryGetValue(key, out function) || m_parseErrors.TryGetValue(errorKey, out parseError))
                {
                    return function != null;
                }
                try
                {
                    function = m_functions.GetOrAdd(key, Parse(returnType, parameterTypes));
                    return true;
                }
                // the parser reports some incomplete scripts (a >) with an exception of the expression tree
                catch (Exception e) when (e is DynamicExpresso.Exceptions.ParseException or ArgumentException or InvalidOperationException)
                {
                    parseError = m_parseErrors.GetOrAdd(errorKey, e.Message);
                    return false;
                }
            }

            private Func<object[], object> Parse(Type returnType, Type[] parameterTypes)
            {
                var parameters = Bindings.Select((x, i) => new Parameter(x.Name, parameterTypes[i])).ToArray();
                var lambda = SharedInterpreter.Parse(Text, returnType, parameters);

                // (object[] args) => { T0 p0 = (T0)args[0]; ...; return (object)body; } - called without reflection,
                // so exceptions of the model (e.g. a binding error) are not wrapped in a TargetInvocationException
                var args = Expression.Parameter(typeof(object[]), "args");
                var variables = parameters.Select(x => x.Expression).ToList();
                var body = variables
                    .Select((x, i) => (Expression)Expression.Assign(x, Expression.Convert(Expression.ArrayIndex(args, Expression.Constant(i)), x.Type)))
                    .Append(Expression.Convert(lambda.Expression, typeof(object)));
                return Expression.Lambda<Func<object[], object>>(Expression.Block(variables, body), args).Compile();
            }
        }

        private readonly struct SignatureKey : IEquatable<SignatureKey>
        {
            private readonly Type m_returnType;
            private readonly Type[] m_parameterTypes;

            public SignatureKey(Type returnType, Type[] parameterTypes)
            {
                m_returnType = returnType;
                m_parameterTypes = parameterTypes;
            }

            public bool Equals(SignatureKey other)
            {
                return m_returnType == other.m_returnType && m_parameterTypes.SequenceEqual(other.m_parameterTypes);
            }

            public override bool Equals(object obj)
            {
                return obj is SignatureKey other && Equals(other);
            }

            public override int GetHashCode()
            {
                int hash = m_returnType.GetHashCode();
                foreach (var type in m_parameterTypes)
                {
                    hash = unchecked((hash * 31) + type.GetHashCode());
                }
                return hash;
            }
        }
    }
}
