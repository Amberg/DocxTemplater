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
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance;
            var type = GetType();
            // Properties first, then public fields - the same precedence as the default model lookup.
            // DisplayNameAttribute cannot be applied to fields, so fields are only found by their name.
            var members = type.GetProperties(flags).Cast<MemberInfo>().Concat(type.GetFields(flags)).ToList();

            // Try to find property with the specified display name.
            foreach (var member in members)
            {
                var displayNameAttribute = member.GetCustomAttribute<DisplayNameAttribute>();
                if (displayNameAttribute is null || !string.Equals(displayNameAttribute.DisplayName, propertyName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return GetValue(member, out value);
            }

            // Try to find property or field with the specified name (case-insensitive) as fallback.
            var memberByName = members.FirstOrDefault(m => string.Equals(m.Name, propertyName, StringComparison.OrdinalIgnoreCase));
            if (memberByName != null)
            {
                return GetValue(memberByName, out value);
            }

            value = new ValueWithMetadata(null);
            return false;
        }

        private bool GetValue(MemberInfo member, out ValueWithMetadata value)
        {
            var modelPropertyAttribute = member.GetCustomAttribute<ModelPropertyAttribute>();

            value = new ValueWithMetadata(ModelLookup.GetMemberValue(member, this), new ValueMetadata(modelPropertyAttribute?.DefaultFormatter));
            return true;
        }
    }
}
