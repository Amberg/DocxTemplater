﻿using DynamicExpresso;
using System;
using System.Dynamic;
using System.Text.RegularExpressions;

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
            scriptAsString = scriptAsString.Trim().Replace('\'', '"').Replace('“', '"');
            // replace leading dots (implicit scope) with variables
            var interpreter = new Interpreter();
            try
            {
                scriptAsString = RegexWordStartingWithDot.Replace(scriptAsString, (m) => OnVariableReplace(m, interpreter));
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
            string prop = match.Groups["prop"].Value;   // Preserve the word after dots
            var dots = match.Groups["dots"].Value;

            var dotCount = dots.Length;
            var scope = m_modelDictionary.GetScopeParentLevel(dotCount - 1);
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

            // If you try to get a value of a property
            // not defined in the class, this method is called.
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

            // If you try to set a value of a property that is
            // not defined in the class, this method is called.
            public override bool TrySetMember(SetMemberBinder binder, object value)
            {
                return false;
            }

        }

    }
}
