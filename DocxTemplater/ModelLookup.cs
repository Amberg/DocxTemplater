﻿using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

namespace DocxTemplater
{
    internal class ModelLookup
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
                    return model;
                }
            }
            if (model == null)
            {
                throw new OpenXmlTemplateException($"Model {variableName} not found");
            }

            for (int i = partIndex; i < parts.Length; i++)
            {
                var propertyName = parts[i];
                if (model is ITemplateModel templateModel)
                {
                    if (!templateModel.TryGetPropertyValue(propertyName, out model))
                    {
                        throw new OpenXmlTemplateException($"Property {propertyName} not found in {modelRootPath}");
                    }
                }
                else if (model is IDictionary<string, object> dict)
                {
                    if (!dict.TryGetValue(parts[i], out model))
                    {
                        throw new OpenXmlTemplateException($"Property {propertyName} not found in {modelRootPath}");
                    }
                }
                else
                {
                    var property = model.GetType().GetProperty(propertyName, BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance);
                    if (property != null)
                    {
                        model = property.GetValue(model);
                        if (model == null)
                        {
                            // if a property is null, we can't continue searching
                            //same behavior as null propagation in C#
                            // ae A.B.C.D --> A?.B?.C?.D
                            return null;
                        }
                    }
                    else if (model is ICollection)
                    {
                        throw new OpenXmlTemplateException($"Property {propertyName} on collection {modelRootPath} not found - is collection start missing? '#{variableName}'");
                    }
                    else
                    {
                        throw new OpenXmlTemplateException($"Property {propertyName} not found in {modelRootPath}");
                    }
                }
            }
            return model;
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
