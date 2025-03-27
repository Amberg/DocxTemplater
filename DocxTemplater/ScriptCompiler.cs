using DynamicExpresso;
using System;
using System.Dynamic;
using System.Text.RegularExpressions;
using DocxTemplater.Model;

namespace DocxTemplater
{
    internal class ScriptCompiler : IScriptCompiler
    {
        private readonly IModelLookup m_modelDictionary;

        private static readonly Regex RegexWordStartingWithDot = new(@"
                                                                    (?x)                          # Enable verbose mode - allows comments and whitespace in pattern
                                                                    (?:                           # Non-capturing group for all possible prefixes:
                                                                        ^                         # Either start of string
                                                                        |                        
                                                                        (?<unary>[+\-!])         # Capture unary operators (+, -, !) in 'unary' group
                                                                        |                       
                                                                        (?<=[^.\p{L}\p{N}_])     # Position after any char that's not a dot, letter, number, or underscore
                                                                                                 # (lookbehind - ensures we don't break existing identifiers)
                                                                    )
                                                                    (?<dots>\.+)                 # Capture one or more dots in 'dots' group
                                                                    (?<prop>                     # Start 'prop' group for the property name
                                                                        [\p{L}\p{N}_]*           # Any number of letters, numbers, or underscores
                                                                    )                            
                                                                    ",
            RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace, TimeSpan.FromMilliseconds(500)
        );

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
                scriptAsString =
                    RegexWordStartingWithDot.Replace(scriptAsString, (m) => OnVariableReplace(m, interpreter));
            }
            catch (RegexMatchTimeoutException)
            {
                throw new OpenXmlTemplateException($"Invalid expression '{scriptAsString}'");
            }

            var identifiers = interpreter.DetectIdentifiers(scriptAsString);
            foreach (var identifier in identifiers.UnknownIdentifiers)
            {
                var val = m_modelDictionary.GetValue(identifier);
                if (val == null || val.GetType().IsPrimitive)
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
            var scope = m_modelDictionary.GetScopeParentLevel(dotCount - 1);

            if (scope != null && !scope.GetType().IsPrimitive && scope is not string)
            {
                scope = new TemplateModelWrapper(scope);
            }

            var varName = $"{unary}__s{dotCount}_"; // choose a variable name that is unlikely to be used by the user
            interpreter.SetVariable(varName, scope);
            if (!string.IsNullOrWhiteSpace(prop))
            {
                varName += $".{prop}";
            }

            return varName;
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
                if (result != null && !result.GetType().IsPrimitive && result is not string)
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

        /// <summary>
        /// Wrapper to handle ITemplatemodels in dynamic expressions
        /// </summary>
        private class TemplateModelWrapper : DynamicObject
        {
            private readonly object m_wrappedModel;

            public TemplateModelWrapper(object wrappedModel)
            {
                m_wrappedModel = wrappedModel;
            }

            public override bool TryGetMember(GetMemberBinder binder, out object result)
            {
                result = null;
                if (m_wrappedModel is ITemplateModel templateModel)
                {
                    if (templateModel.TryGetPropertyValue(binder.Name, out ValueWithMetadata valueWithMetadata))
                    {
                        result = valueWithMetadata.Value;
                        return true;
                    }
                }
                var prop = m_wrappedModel.GetType().GetProperty(binder.Name);
                if (prop != null)
                {
                    result = prop.GetValue(m_wrappedModel);
                    if (result != null && !result.GetType().IsPrimitive && result is not string)
                    {
                        result = new TemplateModelWrapper(result);
                    }
                    return true;
                }
                return false;
            }
            public override bool TrySetMember(SetMemberBinder binder, object value)
            {
                return false;
            }
        }

    }
}
