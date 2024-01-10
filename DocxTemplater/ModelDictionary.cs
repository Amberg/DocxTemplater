using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace DocxTemplater
{
    internal class ModelDictionary
    {
        private readonly Dictionary<string, object> m_models;
        private readonly Dictionary<string, object> m_loopVariables;
        private readonly Lazy<string> m_defaultModelPrefix;

        private string m_rootModelPrefix;

        public ModelDictionary()
        {
            m_models = new Dictionary<string, object>();
            m_loopVariables = new Dictionary<string, object>();
            m_defaultModelPrefix = new Lazy<string>(() => m_rootModelPrefix = m_models.Keys.FirstOrDefault());
        }

        public IReadOnlyDictionary<string, object> Models => m_models;

        public void Add(string prefix, object model)
        {
            m_models.Add(prefix, model);
        }

        public void AddLoopVariable(string name, object value)
        {
            name = AddPathPrefixInSingleModelMode(name);
            m_loopVariables.Add(name, value);
        }

        public bool IsLoopVariable(string name)
        {
            name = AddPathPrefixInSingleModelMode(name);
            return m_loopVariables.ContainsKey(name);
        }

        public void RemoveLoopVariable(string name)
        {
            name = AddPathPrefixInSingleModelMode(name);
            m_loopVariables.Remove(name);
        }

        private string AddPathPrefixInSingleModelMode(string name)
        {
            var dotIndex = name.IndexOf('.');
            if (dotIndex == -1 || !m_models.ContainsKey(name[..dotIndex]))
            {
                if (m_defaultModelPrefix.Value != null &&
                    !name.Equals(m_rootModelPrefix, StringComparison.CurrentCultureIgnoreCase) &&
                    !name.StartsWith(m_defaultModelPrefix.Value + ".", StringComparison.CurrentCultureIgnoreCase))
                {
                    name = $"{m_rootModelPrefix}.{name}";
                }
            }

            return name;
        }

        public object GetValue(string variableName)
        {
            var parts = variableName.Split('.');
            var path = parts[0];

            int startIndex = 0;
            if (!m_models.ContainsKey(path))
            {
                startIndex = -1;
                path = m_defaultModelPrefix.Value;
            }
            object model = null;
            for (int i = startIndex; i < parts.Length; i++)
            {
                if (!m_loopVariables.TryGetValue(path, out var nextModel) && !m_models.TryGetValue(path, out nextModel))
                {
                    if (model == null)
                    {
                        throw new OpenXmlTemplateException($"Model {path} not found");
                    }
                    var property = model.GetType().GetProperty(parts[i], BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance);
                    if (property != null)
                    {
                        model = property.GetValue(model);
                    }
                    else if (model is ICollection)
                    {
                        throw new OpenXmlTemplateException($"Property {parts[i]} on collection {path} not found - is collection start missing? '#{variableName}'");
                    }
                    else
                    {
                        throw new OpenXmlTemplateException($"Property {parts[i]} not found in {path}");
                    }
                }
                else
                {
                    model = nextModel;
                }
                if (i + 1 < parts.Length)
                {
                    path = $"{path}.{parts[i + 1]}";
                }
            }
            return model;
        }
    }
}
