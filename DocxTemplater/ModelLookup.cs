using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace DocxTemplater
{
    internal class ModelLookup : IModelLookup
    {
        private readonly Dictionary<string, object> m_rootScope;
        private readonly Stack<Dictionary<string, object>> m_blockScopes;

        public ModelLookup()
        {
            m_rootScope = new Dictionary<string, object>();
            m_blockScopes = new Stack<Dictionary<string, object>>();
            m_blockScopes.Push(m_rootScope);
        }

        public IReadOnlyDictionary<string, object> Models => m_rootScope;

        public void Add(string prefix, object model)
        {
            m_rootScope.Add(prefix, model);
        }

        public IVariableScope OpenScope()
        {
            return new VariableScope(m_blockScopes);
        }


        public object GetValue(string variableName)
        {
            return GetValueWithMetadata(variableName).Value;
        }

        public ValueWithMetadata GetValueWithMetadata(string variableName)
        {
            var leadingDotsCount = variableName.TakeWhile(x => x == '.').Count();
            variableName = variableName[leadingDotsCount..];
            int partIndex = 0;
            var parts = variableName.Split('.');
            object model = null;
            string modelRootPath = variableName;
            if (leadingDotsCount == 0)
            {
                model = SearchLongestPathInLookup(parts, out modelRootPath, out partIndex, 0);
                if (model == null && m_rootScope.Count > 0)
                {
                    var firstModelEntry = m_rootScope.First();
                    // a.b.c.d and b.c.d.e ==> a.b.c.d.e
                    parts = firstModelEntry.Key.Split('.').Concat(variableName.Split('.')).Distinct().ToArray();
                    model = SearchLongestPathInLookup(parts, out modelRootPath, out partIndex, 0);
                }

                if (model == null)
                {
                    throw new OpenXmlTemplateException($"Model {variableName} not found");
                }
            }
            else
            {
                modelRootPath = "parent scope";
                model = m_blockScopes.ElementAt(leadingDotsCount - 1).Values.FirstOrDefault();
                if (parts.Length == 1 && string.IsNullOrWhiteSpace(parts[0]))
                {
                    return new ValueWithMetadata(model, new ValueMetadata());
                }
            }
            if (model == null)
            {
                throw new OpenXmlTemplateException($"Model {variableName} not found");
            }

            PropertyInfo lastProperty = null;
            ValueMetadata? lastValueMetadata = null;
            for (int i = partIndex; i < parts.Length; i++)
            {
                lastValueMetadata = null;
                var propertyName = parts[i];
                if (model is ITemplateModel templateModel)
                {
                    if (!templateModel.TryGetPropertyValue(propertyName, out ValueWithMetadata valWithMetadata))
                    {
                        throw new OpenXmlTemplateException($"Property {propertyName} not found in {modelRootPath}");
                    }
                    model = valWithMetadata.Value;
                    lastValueMetadata = valWithMetadata.Metadata;
                }
                else if (model is IDictionary<string, object> dict)
                {
                    if (!dict.TryGetValue(propertyName, out model))
                    {
                        throw new OpenXmlTemplateException($"Property {propertyName} not found in {modelRootPath}");
                    }
                    if (model is ValueWithMetadata valWithMetadata)
                    {
                        model = valWithMetadata.Value;
                        lastValueMetadata = valWithMetadata.Metadata;
                    }
                }
                else
                {
                    var property = model.GetType().GetProperty(propertyName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance);
                    lastProperty = property;
                    if (property != null)
                    {
                        model = property.GetValue(model);
                        if (model == null)
                        {
                            // if a property is null, we can't continue searching
                            //same behavior as null propagation in C#
                            // ae A.B.C.D --> A?.B?.C?.D
                            return new ValueWithMetadata(null, new ValueMetadata());
                        }
                    }
                    else if (model is ICollection)
                    {
                        throw new OpenXmlTemplateException($"Property '{variableName}' on collection of type '{model.GetType()}' not found");
                    }
                    else
                    {
                        throw new OpenXmlTemplateException($"Property '{propertyName}' not found in '{variableName}' of type '{model.GetType()}'");
                    }
                }
            }

            if (lastProperty != null)
            {
                var metadata = lastProperty.GetCustomAttribute<ModelPropertyAttribute>();
                return new ValueWithMetadata(model, new ValueMetadata(metadata?.DefaultFormatter));
            }
            return new ValueWithMetadata(model, lastValueMetadata ?? new ValueMetadata());
        }

        private object SearchLongestPathInLookup(string[] parts, out string modelRootPath, out int partIndex, int startScopeIndex)
        {
            modelRootPath = null;
            partIndex = parts.Length;
            foreach (Dictionary<string, object> scope in m_blockScopes.Skip(startScopeIndex))
            {
                partIndex = parts.Length;
                // search the longest path in the lookup
                for (; partIndex > 0; partIndex--)
                {
                    modelRootPath = string.Join('.', parts[..partIndex]);
                    if (scope.TryGetValue(modelRootPath, out var model))
                    {
                        return model;
                    }
                }
            }
            return null;
        }


        internal class VariableScope : IVariableScope
        {
            private readonly Dictionary<string, object> m_scope;
            private readonly Stack<Dictionary<string, object>> m_scopeStack;

            public VariableScope(Stack<Dictionary<string, object>> scopeStack)
            {
                m_scopeStack = scopeStack;
                m_scope = new Dictionary<string, object>();
                scopeStack.Push(m_scope);
            }

            public void AddVariable(string name, object value)
            {
                // remove leading dots
                name = name.TrimStart('.');
                Debug.Assert(m_scopeStack.Count > 1, "Added Block variable in root scope");
                m_scope.Add(name, value);
            }

            public void Dispose()
            {
                m_scopeStack.Pop();
            }
        }

        public object GetScopeParentLevel(int parentLevel)
        {
            return m_blockScopes.ElementAt(parentLevel).Values.FirstOrDefault();
        }
    }
}
