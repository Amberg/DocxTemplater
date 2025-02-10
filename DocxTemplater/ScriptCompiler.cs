using DynamicExpresso;
using System;
using System.Dynamic;
using System.Text.RegularExpressions;

namespace DocxTemplater
{
    internal class ScriptCompiler : IScriptCompiler
    {
        private readonly IModelLookup m_modelDictionary;
        private static readonly Regex RegexWordStartingWithDot = new(@"(?:^|\s+)(\.+)([\p{L}\p{N}_]*)", RegexOptions.Compiled);

        public ScriptCompiler(IModelLookup modelDictionary, ProcessSettings processSettings)
        {
            this.m_modelDictionary = modelDictionary;
            ProcessSettings = processSettings;
        }

        public ProcessSettings ProcessSettings { get; }

        public Func<bool> CompileScript(string scriptAsString)
        {
            scriptAsString = scriptAsString.Trim().Replace('\'', '"');

            // replace replace leading dots (implicit scope) with variables
            var interpreter = new Interpreter();
            scriptAsString = RegexWordStartingWithDot.Replace(scriptAsString, (m) => OnVariableReplace(m, interpreter));
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
            var dotCount = match.Groups[1].Length;
            var scope = m_modelDictionary.GetScopeParentLevel(dotCount - 1);
            var varName = $"__s{dotCount}_"; // choose a variable name that is unlikely to be used by the user
            interpreter.SetVariable(varName, scope);
            if (!string.IsNullOrWhiteSpace(match.Groups[2].Value))
            {
	            varName += $".{match.Groups[2].Value}";
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
