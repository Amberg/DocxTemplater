using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenXml.Templates
{
    internal class ModelDictionary
    {
        private readonly Dictionary<string, object> m_models;
        private string m_rootModelPrefix;

        public ModelDictionary()
        {
            m_models = new Dictionary<string, object>();
        }

        public void Add(string prefix, object model)
        {
            if (m_rootModelPrefix != null && !prefix.StartsWith(m_rootModelPrefix))
            {
                prefix = $"{m_rootModelPrefix}.{prefix}";
            }
            m_models.Add(prefix, model);
        }

        public void SetModelPrefix()
        {
            m_rootModelPrefix = m_models.Count == 1 ? m_models.Keys.First() : null;
        }


        public object GetValue(string variableName)
        {
            if (m_rootModelPrefix != null && !variableName.StartsWith(m_rootModelPrefix))
            {
                variableName = $"{m_rootModelPrefix}.{variableName}";
            }
            var parts = variableName.Split('.');
            var path = parts[0];
            object model = null;
            for (int i = 0; i < parts.Length; i++)
            {
                if (!m_models.TryGetValue(path, out var nextModel))
                {
                    if (model == null)
                    {
                        throw new OpenXmlTemplateException($"Model {path} not found");
                    }
                    var property = model.GetType().GetProperty(parts[i]);
                    if (property != null)
                    {
                        model = property.GetValue(model);
                    }
                    else if(model is ICollection)
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

        public void Remove(string name)
        {
            if (m_rootModelPrefix != null && !name.StartsWith(m_rootModelPrefix))
            {
                name = $"{m_rootModelPrefix}.{name}";
            }
            m_models.Remove(name);
        }
    }
}
