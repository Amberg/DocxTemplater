using DynamicExpresso;
using System;
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
                                                                        (?<=[^.\p{L}\p{N}_?])    # Position after any char that's not a dot, letter, number, underscore, or question mark (for null-conditional)
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

        public ScriptCompiler(IModelLookup modelDictionary, ProcessSettings processSettings)
        {
            this.m_modelDictionary = modelDictionary;
            ProcessSettings = processSettings;
        }

        public ProcessSettings ProcessSettings { get; }

        public Func<bool> CompileScript(string scriptAsString)
        {
            var interpreter = CreateInterpreter(ref scriptAsString);

            try
            {
                var func = interpreter.ParseAsDelegate<Func<bool>>(scriptAsString);
                return WrapExecution(func);
            }
            catch (DynamicExpresso.Exceptions.ParseException e)
            {
                throw new OpenXmlTemplateException($"Error parsing script {scriptAsString}", e);
            }
        }

        public Func<object> CompileExpression(string scriptAsString)
        {
            var interpreter = CreateInterpreter(ref scriptAsString);

            try
            {
                var func = interpreter.ParseAsDelegate<Func<object>>(scriptAsString);
                return WrapExecution(func);
            }
            catch (DynamicExpresso.Exceptions.ParseException e)
            {
                throw new OpenXmlTemplateException($"Error parsing expression {scriptAsString}", e);
            }
        }

        private static Func<T> WrapExecution<T>(Func<T> func)
        {
            return () =>
            {
                try
                {
                    return func();
                }
                catch (Exception e) when (e is NullReferenceException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or System.Reflection.TargetInvocationException)
                {
                    var message = e is System.Reflection.TargetInvocationException tie ? tie.InnerException?.Message ?? e.Message : e.Message;
                    throw new OpenXmlTemplateException(message, e);
                }
            };
        }

        private Interpreter CreateInterpreter(ref string scriptAsString)
        {
            scriptAsString = HelperFunctions.SanitizeQuotes(scriptAsString);
            // replace leading dots (implicit scope) with variables
            var interpreter = new Interpreter().EnableAssignment(AssignmentOperators.None);
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
                object val;
                try
                {
                    val = m_modelDictionary.GetValue(identifier);
                }
                catch (OpenXmlTemplateException) when (ProcessSettings?.BindingErrorHandling is BindingErrorHandling.SkipBindingAndRemoveContent)
                {
                    val = null;
                }
                if (val == null || IsSimpleType(val.GetType()))
                {
                    interpreter.SetVariable(identifier, val, val?.GetType() ?? typeof(string));
                }
                else
                {
                    interpreter.SetVariable(identifier, new ScriptCompilerModelVariable(m_modelDictionary, identifier, ProcessSettings));
                }
            }

            return interpreter;
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
                    scope = new ScriptCompilerModelVariable(m_modelDictionary, new string('.', dotCount - 1), ProcessSettings);
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
    }
}
