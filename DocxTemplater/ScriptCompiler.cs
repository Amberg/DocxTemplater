using System;
using System.Text.RegularExpressions;

namespace DocxTemplater
{
    /// <summary>
    /// Compiles the conditions and expressions of a template. The parsed scripts are kept in a <see cref="ScriptCache"/>:
    /// the model values are passed as parameters, so a script is parsed once per combination of value types and
    /// only evaluated after that.
    /// </summary>
    internal class ScriptCompiler : IScriptCompiler
    {
        private readonly IModelLookup m_modelDictionary;
        private readonly ScriptCache m_cache;

        /// <param name="processSettings">Holds the <see cref="ScriptCache"/> shared by all templates created with them.</param>
        public ScriptCompiler(IModelLookup modelDictionary, ProcessSettings processSettings)
        {
            m_modelDictionary = modelDictionary;
            ProcessSettings = processSettings;
            m_cache = processSettings.ScriptCache;
        }

        public ProcessSettings ProcessSettings { get; }

        public Func<bool> CompileScript(string scriptAsString)
        {
            var compiled = Compile(scriptAsString, typeof(bool), TemplateErrorCode.ScriptParseError);
            return () => (bool)compiled();
        }

        public Func<object> CompileExpression(string scriptAsString)
        {
            return Compile(scriptAsString, typeof(object), TemplateErrorCode.ExpressionParseError);
        }

        private Func<object> Compile(string scriptAsString, Type returnType, TemplateErrorCode parseError)
        {
            ScriptCache.PreparedScript prepared;
            try
            {
                prepared = m_cache.Prepare(scriptAsString);
            }
            catch (RegexMatchTimeoutException)
            {
                throw OpenXmlTemplateException.Create(ProcessSettings, TemplateErrorCode.InvalidExpression, HelperFunctions.SanitizeQuotes(scriptAsString));
            }

            var bindings = prepared.Bindings;
            var values = new object[bindings.Count];
            var types = new Type[bindings.Count];
            for (int i = 0; i < bindings.Count; i++)
            {
                values[i] = ResolveBinding(bindings[i], out types[i]);
            }

            if (!prepared.TryGetFunction(returnType, types, out var func, out var parserMessage))
            {
                throw OpenXmlTemplateException.Create(ProcessSettings, new TemplateException(parserMessage), parseError, prepared.Text);
            }
            return WrapExecution(() => func(values));
        }

        /// <param name="type">The type the script is parsed with: the type of the value, a fixed type for <c>null</c>.</param>
        private object ResolveBinding(ScriptCache.Binding binding, out Type type)
        {
            object value;
            switch (binding.Kind)
            {
                case ScriptCache.BindingKind.Identifier:
                    try
                    {
                        value = m_modelDictionary.GetValue(binding.Path);
                    }
                    catch (OpenXmlTemplateException) when (ProcessSettings?.BindingErrorHandling is BindingErrorHandling.SkipBindingAndRemoveContent)
                    {
                        value = null;
                    }
                    if (value == null || IsSimpleType(value.GetType()))
                    {
                        type = value?.GetType() ?? typeof(string);
                        return value;
                    }
                    value = new ScriptCompilerModelVariable(m_modelDictionary, binding.Path, ProcessSettings);
                    break;
                case ScriptCache.BindingKind.Scope:
                    value = m_modelDictionary.GetScopeParentLevel(binding.DotCount - 1);
                    if (value != null && !IsSimpleType(value.GetType()))
                    {
                        value = new ScriptCompilerModelVariable(m_modelDictionary, new string('.', binding.DotCount - 1), ProcessSettings);
                    }
                    break;
                default:
                    value = m_modelDictionary.GetValue(binding.Path);
                    break;
            }
            type = value?.GetType() ?? typeof(object);
            return value;
        }

        private static Func<T> WrapExecution<T>(Func<T> func)
        {
            return () =>
            {
                try
                {
                    return func();
                }
                catch (Exception e) when (e is NullReferenceException or Microsoft.CSharp.RuntimeBinder.RuntimeBinderException or System.Reflection.TargetInvocationException)
                {
                    var message = e is System.Reflection.TargetInvocationException tie ? tie.InnerException?.Message ?? e.Message : e.Message;
                    throw new OpenXmlTemplateException(message, e);
                }
            };
        }

        /// <summary>
        /// Determines whether the specified type is a simple type.
        /// Simple types include primitive types, enums, strings, decimals, DateTime, and GUIDs.
        /// </summary>
        internal static bool IsSimpleType(Type type)
        {
            var typeCode = Type.GetTypeCode(type);
            return typeCode != TypeCode.Object || type == typeof(Guid);
        }
    }
}
