using DynamicExpresso;
using System;
using System.Dynamic;

namespace DocxTemplater
{
    internal class ScriptCompiler
    {
        private readonly ModelDictionary m_modelDictionary;

        public ScriptCompiler(ModelDictionary modelDictionary)
        {
            this.m_modelDictionary = modelDictionary;
        }

        public Func<bool> CompileScript(string scriptAsString)
        {
            var interpreter = new Interpreter();
            var identifiers = interpreter.DetectIdentifiers(scriptAsString);
            foreach (var identifier in identifiers.UnknownIdentifiers)
            {
                interpreter.SetVariable(identifier, new ModelVariable(m_modelDictionary, identifier));
            }
            return interpreter.ParseAsDelegate<Func<bool>>(scriptAsString);
        }

        private class ModelVariable : DynamicObject
        {
            private readonly ModelDictionary m_modelDictionary;
            private readonly string m_rootName;

            public ModelVariable(ModelDictionary modelDictionary, string rootName)
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
                if (m_modelDictionary.IsLoopVariable(name))
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
