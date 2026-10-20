using System;
using System.Collections.Generic;
using System.Globalization;

namespace DocxTemplater
{
    /// <summary>
    /// Base class of the errors raised by the template engine, independent of the template type
    /// (<see cref="OpenXmlTemplateException"/> for Word documents). Carries the <see cref="ErrorCode"/> and the
    /// <see cref="Arguments"/> of the message, so an application can react to an error - or translate it - without
    /// parsing the message text.
    /// </summary>
    [Serializable]
    public class TemplateException : Exception
    {
        // the provider the message was formatted with, so GetMessage(culture) uses the same languages
        private readonly TemplateErrorMessages m_messages;

        /// <summary>
        /// Creates an exception with a free text message (<see cref="ErrorCode"/> is <see cref="TemplateErrorCode.None"/>).
        /// </summary>
        public TemplateException(string message) : base(message)
        {
            Arguments = Array.Empty<object>();
        }

        /// <inheritdoc cref="TemplateException(string)"/>
        public TemplateException(string message, Exception inner) : base(message, inner)
        {
            Arguments = Array.Empty<object>();
        }

        /// <summary>
        /// Creates an exception for <paramref name="errorCode"/> with an already formatted message.
        /// Use <see cref="Create(ProcessSettings, TemplateErrorCode, object[])"/> to format the message in the
        /// language configured in the <see cref="ProcessSettings"/>.
        /// </summary>
        public TemplateException(TemplateErrorCode errorCode, IReadOnlyList<object> arguments, string message, Exception inner = null)
            : this(errorCode, arguments, message, inner, null)
        {
        }

        /// <param name="messages">The provider the <paramref name="message"/> was formatted with; <c>null</c> for <see cref="TemplateErrorMessages.Default"/>.</param>
        protected TemplateException(TemplateErrorCode errorCode, IReadOnlyList<object> arguments, string message, Exception inner, TemplateErrorMessages messages)
            : base(message, inner)
        {
            ErrorCode = errorCode;
            Arguments = arguments ?? Array.Empty<object>();
            m_messages = messages;
        }

        /// <summary>
        /// Creates an exception whose message is formatted with <see cref="ProcessSettings.ErrorMessages"/> in the
        /// <see cref="ProcessSettings.UiCulture"/>. <paramref name="settings"/> may be <c>null</c> (English).
        /// </summary>
        public static TemplateException Create(ProcessSettings settings, TemplateErrorCode errorCode, params object[] arguments)
        {
            return Create(settings, null, errorCode, arguments);
        }

        /// <inheritdoc cref="Create(ProcessSettings, TemplateErrorCode, object[])"/>
        public static TemplateException Create(ProcessSettings settings, Exception innerException, TemplateErrorCode errorCode, params object[] arguments)
        {
            arguments ??= Array.Empty<object>();
            var message = FormatMessage(settings, errorCode, arguments, out var messages);
            return new TemplateException(errorCode, arguments, message, innerException, messages);
        }

        /// <summary>
        /// Formats the message of <paramref name="errorCode"/> with <see cref="ProcessSettings.ErrorMessages"/> in the
        /// <see cref="ProcessSettings.UiCulture"/> (English if <paramref name="settings"/> is <c>null</c>), for the
        /// factory methods of derived classes.
        /// </summary>
        /// <param name="messages">The provider the message was formatted with, to pass on to the constructor.</param>
        protected static string FormatMessage(ProcessSettings settings, TemplateErrorCode errorCode, IReadOnlyList<object> arguments, out TemplateErrorMessages messages)
        {
            messages = settings?.ErrorMessages ?? TemplateErrorMessages.Default;
            var culture = settings?.UiCulture ?? CultureInfo.InvariantCulture;
            return messages.Format(errorCode, culture, arguments ?? Array.Empty<object>());
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
        /// The message in another language than the one the exception was created with, using the
        /// <see cref="TemplateErrorMessages"/> the exception was created with (<see cref="TemplateErrorMessages.Default"/>
        /// if it was not created through <see cref="Create(ProcessSettings, TemplateErrorCode, object[])"/>).
        /// Falls back to <see cref="Exception.Message"/> for free text errors.
        /// </summary>
        public string GetMessage(CultureInfo culture)
        {
            return GetMessage(m_messages ?? TemplateErrorMessages.Default, culture);
        }

        /// <inheritdoc cref="GetMessage(CultureInfo)"/>
        public string GetMessage(TemplateErrorMessages messages, CultureInfo culture)
        {
            ArgumentNullException.ThrowIfNull(messages);
            return ErrorCode == TemplateErrorCode.None ? Message : messages.Format(ErrorCode, culture, Arguments);
        }
    }
}
