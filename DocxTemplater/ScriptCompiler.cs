using DynamicExpresso;
using System;
using System.Dynamic;
using System.Text.RegularExpressions;
using DocxTemplater.Blocks;

namespace DocxTemplater
{
    internal class ScriptCompiler : IScriptCompiler
    {
        private readonly IModelLookup m_modelDictionary;

        private static readonly Regex RegexWordStartingWithDot = new(@"
                                                                    (?x)                         # Enable verbose mode - allows comments and whitespace in pattern
                                                                    (?:                          # Non-capturing group for all possible prefixes:
                                                                        ^                        # Either start of string
                                                                        |                        
                                                                        (?<unary>[+\-!])         # Capture unary operators (+, -, !) in 'unary' group
                                                                        |                       
                                                                        (?<=[^.\p{L}\p{N}_])     # Position after any char that's not a dot, letter, number, or underscore
                                                                                                 # (lookbehind - ensures we don't break existing identifiers)

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

        public ScriptCompiler(IModelLookup modelDictionary, ProcessSettings processSettings)
        {
            this.m_modelDictionary = modelDictionary;
            ProcessSettings = processSettings;
        }

        public ProcessSettings ProcessSettings { get; }

        public Func<bool> CompileScript(string scriptAsString)
        {
            scriptAsString = HelperFunctions.SanitizeQuotes(scriptAsString);
            // replace leading dots (implicit scope) with variables
            var interpreter = new Interpreter();
            try
            {
                // replace ..foo / .foo etc
                scriptAsString = RegexWordStartingWithDot.Replace(scriptAsString, (m) => OnVariableReplace(m, interpreter));

                scriptAsString = RegexListKeywords.Replace(scriptAsString, (m) => OnSpecialKeyWordsReplace(m, interpreter));

            }
            catch (RegexMatchTimeoutException)
            {
                throw new OpenXmlTemplateException($"Invalid expression '{scriptAsString}'");
            }

            var identifiers = interpreter.DetectIdentifiers(scriptAsString);
            foreach (var identifier in identifiers.UnknownIdentifiers)
            {
                var val = m_modelDictionary.GetValue(identifier);
                if (val == null || IsSimpleType(val.GetType()))
                {
                    interpreter.SetVariable(identifier, val);
                }
                else
                {
                    interpreter.SetVariable(identifier, new ModelVariable(m_modelDictionary, identifier));
                }
            }

            try
            {
                return interpreter.ParseAsDelegate<Func<bool>>(scriptAsString);
            }
            catch (DynamicExpresso.Exceptions.ParseException e)
            {
                throw new OpenXmlTemplateException($"Error parsing script {scriptAsString}", e);
            }
        }

        private string OnVariableReplace(Match match, Interpreter interpreter)
        {
            string unary = match.Groups["unary"].Value; // Preserve unary operators if present
            string prop = match.Groups["prop"].Value; // Preserve the word after dots
            var dots = match.Groups["dots"].Value;
            var dotCount = dots.Length;

            if (prop == LoopBlock.LoopIndexVariable || prop == LoopBlock.LoopLengthVariable)
            {
                var varName = $"__s{dotCount}_{prop}";
                var value = m_modelDictionary.GetValue(match.Value);
                interpreter.SetVariable(varName, value);
                return varName;
            }
            else
            {
                // ".Foo" or "..Foo" etc
                var scope = m_modelDictionary.GetScopeParentLevel(dotCount - 1);
                if (scope != null && !IsSimpleType(scope.GetType()))
                {
                    scope = new ModelVariable(m_modelDictionary, new string('.', dotCount - 1));
                }

                var varName = $"__s{dotCount}_"; // choose a variable name that is unlikely to be used by the user
                interpreter.SetVariable(varName, scope);
                if (!string.IsNullOrWhiteSpace(prop))
                {
                    varName += $".{prop}";
                }

                return $"{unary}{varName}";
            }
        }

        private string OnSpecialKeyWordsReplace(Match match, Interpreter interpreter)
        {
            var dots = match.Groups["dots"].Value.Length;
            var keyWord = match.Groups["keyword"];
            var varName = $"__s{dots}_{keyWord}";
            var value = m_modelDictionary.GetValue(match.Value);
            interpreter.SetVariable(varName, value);
            return varName;
        }

        /// <summary>
        /// Determines whether the specified type is a simple type.
        /// Simple types include primitive types, enums, strings, decimals, DateTime, and GUIDs.
        /// </summary>
        private static bool IsSimpleType(Type type)
        {
            var typeCode = Type.GetTypeCode(type);
            return typeCode != TypeCode.Object || type == typeof(Guid);
        }


        private class ModelVariable : DynamicObject
        {
            private readonly IModelLookup m_modelDictionary;
            private readonly string m_rootName;

            public ModelVariable(IModelLookup modelDictionary, string rootName)
            {
                m_modelDictionary = modelDictionary;
                m_rootName = rootName;
            }

            public override bool TryGetMember(GetMemberBinder binder, out object result)
            {
                var name = m_rootName + "." + binder.Name;
                result = m_modelDictionary.GetValue(name);
                if (result != null && !IsSimpleType(result.GetType()))
                {
                    result = new ModelVariable(m_modelDictionary, name);
                }

                return true;
            }

            public override bool TrySetMember(SetMemberBinder binder, object value)
            {
                return false;
            }
        }
    }
}
