using System;
using System.Collections.Generic;
using System.Globalization;

namespace DocxTemplater
{
    [Serializable]
    public class OpenXmlTemplateException : Exception
    {
        /// <summary>
        /// Creates an exception with a free text message (<see cref="ErrorCode"/> is <see cref="TemplateErrorCode.None"/>).
        /// </summary>
        public OpenXmlTemplateException(string message) : base(message)
        {
            Arguments = Array.Empty<object>();
        }

        /// <inheritdoc cref="OpenXmlTemplateException(string)"/>
        public OpenXmlTemplateException(string message, Exception inner) : base(message, inner)
        {
            Arguments = Array.Empty<object>();
        }

        /// <summary>
        /// Creates an exception for <paramref name="errorCode"/> with an already formatted message.
        /// Use <see cref="Create(ProcessSettings, TemplateErrorCode, object[])"/> to format the message in the
        /// language configured in the <see cref="ProcessSettings"/>.
        /// </summary>
        public OpenXmlTemplateException(TemplateErrorCode errorCode, IReadOnlyList<object> arguments, string message, Exception inner = null)
            : base(message, inner)
        {
            ErrorCode = errorCode;
            Arguments = arguments ?? Array.Empty<object>();
        }

        /// <summary>
        /// Creates an exception whose message is formatted with <see cref="ProcessSettings.ErrorMessages"/> in the
        /// <see cref="ProcessSettings.UiCulture"/>. <paramref name="settings"/> may be <c>null</c> (English).
        /// </summary>
        public static OpenXmlTemplateException Create(ProcessSettings settings, TemplateErrorCode errorCode, params object[] arguments)
        {
            return Create(settings, null, errorCode, arguments);
        }

        /// <inheritdoc cref="Create(ProcessSettings, TemplateErrorCode, object[])"/>
        public static OpenXmlTemplateException Create(ProcessSettings settings, Exception innerException, TemplateErrorCode errorCode, params object[] arguments)
        {
            var messages = settings?.ErrorMessages ?? TemplateErrorMessages.Default;
            var culture = settings?.UiCulture ?? CultureInfo.InvariantCulture;
            arguments ??= Array.Empty<object>();
            return new OpenXmlTemplateException(errorCode, arguments, messages.Format(errorCode, culture, arguments), innerException);
        }

        /// <summary>
        /// Identifies the error independent of the language of the <see cref="Exception.Message"/>.
        /// <see cref="TemplateErrorCode.None"/> for free text messages.
        /// </summary>
        public TemplateErrorCode ErrorCode { get; }

        /// <summary>
        /// The arguments of the message, see the documentation of the <see cref="ErrorCode"/>'s value for their meaning.
        /// </summary>
        public IReadOnlyList<object> Arguments { get; }

        /// <summary>
        /// The message in another language than the one the exception was created with. Falls back to
        /// <see cref="Exception.Message"/> for free text errors.
        /// </summary>
        public string GetMessage(CultureInfo culture)
        {
            return GetMessage(TemplateErrorMessages.Default, culture);
        }

        /// <inheritdoc cref="GetMessage(CultureInfo)"/>
        public string GetMessage(TemplateErrorMessages messages, CultureInfo culture)
        {
            ArgumentNullException.ThrowIfNull(messages);
            return ErrorCode == TemplateErrorCode.None ? Message : messages.Format(ErrorCode, culture, Arguments);
        }
    }
}
