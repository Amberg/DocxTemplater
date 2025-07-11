using System;
using System.ComponentModel;
using System.Linq;
using System.Reflection;

namespace DocxTemplater.Model
{
    public abstract class TemplateModelWithDisplayNames : ITemplateModel
    {
        public bool TryGetPropertyValue(string propertyName, out ValueWithMetadata value)
        {
            var properties = GetType().GetProperties(BindingFlags.IgnoreCase | BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance);

            // Try to find property with the specified display name.
            foreach (var property in properties)
            {
                var displayNameAttribute = property.GetCustomAttribute<DisplayNameAttribute>();
                if (displayNameAttribute is null || !string.Equals(displayNameAttribute.DisplayName, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return GetValue(property, out value);
            }

            // Try to find property with the specified name (case-insensitive) as fallback.
            var propertyByName = properties.FirstOrDefault(p => string.Equals(p.Name, propertyName, StringComparison.OrdinalIgnoreCase));
            if (propertyByName != null)
            {
                return GetValue(propertyByName, out value);
            }

            value = new ValueWithMetadata(null);
            return false;
        }

        private bool GetValue(PropertyInfo property, out ValueWithMetadata value)
        {
            var modelPropertyAttribute = property.GetCustomAttribute<ModelPropertyAttribute>();

            value = new ValueWithMetadata(property.GetValue(this), new ValueMetadata(modelPropertyAttribute?.DefaultFormatter));
            return true;
        }
    }
}
