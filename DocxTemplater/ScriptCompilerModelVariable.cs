using System;
using System.Dynamic;

namespace DocxTemplater
{
    /// <summary>
    /// Class that represents a variable in the script compiler.
    /// It uses dynamic member access to retrieve values from the model dictionary. <see cref="IModelLookup"/>.
    /// It also supports method invocation on the retrieved objects.
    /// </summary>
    internal class ScriptCompilerModelVariable : DynamicObject
    {
        private readonly IModelLookup m_modelDictionary;
        private readonly string m_rootName;
        public ProcessSettings ProcessSettings { get; }

        public ScriptCompilerModelVariable(IModelLookup modelDictionary, string rootName, ProcessSettings processSettings)
        {
            m_modelDictionary = modelDictionary;
            m_rootName = rootName;
            ProcessSettings = processSettings;
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            var name = m_rootName + "." + binder.Name;
            try
            {
                result = m_modelDictionary.GetValue(name);
            }
            catch (OpenXmlTemplateException) when (ProcessSettings?.BindingErrorHandling is BindingErrorHandling.SkipBindingAndRemoveContent)
            {
                result = null;
            }
            if (result != null && !IsSimpleType(result.GetType()))
            {
                result = new ScriptCompilerModelVariable(m_modelDictionary, name, ProcessSettings);
            }

            return true;
        }

        public override bool TryInvokeMember(InvokeMemberBinder binder, object[] args, out object result)
        {
            object actualObject;
            try
            {
                actualObject = m_modelDictionary.GetValue(m_rootName);
            }
            catch (OpenXmlTemplateException) when (ProcessSettings?.BindingErrorHandling is BindingErrorHandling.SkipBindingAndRemoveContent)
            {
                actualObject = null;
            }

            if (actualObject == null)
            {
                result = null;
                return false;
            }

            try
            {
                result = actualObject.GetType().InvokeMember(
                    binder.Name,
                    System.Reflection.BindingFlags.InvokeMethod |
                    System.Reflection.BindingFlags.Instance |
                    System.Reflection.BindingFlags.Public |
                    System.Reflection.BindingFlags.IgnoreCase,
                    null,
                    actualObject,
                    args);
                return true;
            }
            catch (System.MissingMethodException)
            {
                result = null;
                return false;
            }
            catch (Exception e) when (e is System.Reflection.TargetInvocationException or ArgumentException or System.Reflection.TargetParameterCountException or System.Reflection.TargetException)
            {
                var message = e is System.Reflection.TargetInvocationException tie ? tie.InnerException?.Message ?? e.Message : e.Message;
                throw new OpenXmlTemplateException(message, e);
            }
        }
        public override bool TrySetMember(SetMemberBinder binder, object value)
        {
            return false;
        }

        public override string ToString()
        {
            try
            {
                var value = m_modelDictionary.GetValue(m_rootName);
                return value?.ToString() ?? string.Empty;
            }
            catch (OpenXmlTemplateException) when (ProcessSettings?.BindingErrorHandling is BindingErrorHandling.SkipBindingAndRemoveContent)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Determines whether the specified type is a simple type.
        /// Simple types include primitive types, enums, strings, decimals, DateTime, and GUIDs.
        /// </summary>
        private static bool IsSimpleType(Type type)
        {
            var typeCode = Type.GetTypeCode(type);
            return typeCode != TypeCode.Object || type == typeof(Guid);
        }
    }
}
