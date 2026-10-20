using System;
using System.Collections.Generic;

namespace DocxTemplater
{
    /// <summary>
    /// An error raised while rendering a Word document. Derives from <see cref="TemplateException"/>, the common
    /// base of the errors of all template types, which defines <see cref="TemplateException.ErrorCode"/>,
    /// <see cref="TemplateException.Arguments"/> and <see cref="TemplateException.GetMessage(System.Globalization.CultureInfo)"/>.
    /// </summary>
    [Serializable]
    public class OpenXmlTemplateException : TemplateException
    {
        /// <inheritdoc cref="TemplateException(string)"/>
        public OpenXmlTemplateException(string message) : base(message)
        {
        }

        /// <inheritdoc cref="TemplateException(string, Exception)"/>
        public OpenXmlTemplateException(string message, Exception inner) : base(message, inner)
        {
        }

        /// <inheritdoc cref="TemplateException(TemplateErrorCode, IReadOnlyList{object}, string, Exception)"/>
        public OpenXmlTemplateException(TemplateErrorCode errorCode, IReadOnlyList<object> arguments, string message, Exception inner = null)
            : base(errorCode, arguments, message, inner, null)
        {
        }

        private OpenXmlTemplateException(TemplateErrorCode errorCode, IReadOnlyList<object> arguments, string message, Exception inner, TemplateErrorMessages messages)
            : base(errorCode, arguments, message, inner, messages)
        {
        }

        /// <inheritdoc cref="TemplateException.Create(ProcessSettings, TemplateErrorCode, object[])"/>
        public static new OpenXmlTemplateException Create(ProcessSettings settings, TemplateErrorCode errorCode, params object[] arguments)
        {
            return Create(settings, null, errorCode, arguments);
        }

        /// <inheritdoc cref="TemplateException.Create(ProcessSettings, Exception, TemplateErrorCode, object[])"/>
        public static new OpenXmlTemplateException Create(ProcessSettings settings, Exception innerException, TemplateErrorCode errorCode, params object[] arguments)
        {
            arguments ??= Array.Empty<object>();
            var message = FormatMessage(settings, errorCode, arguments, out var messages);
            return new OpenXmlTemplateException(errorCode, arguments, message, innerException, messages);
        }
    }
}
