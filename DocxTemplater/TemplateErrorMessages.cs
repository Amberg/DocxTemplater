using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace DocxTemplater
{
    /// <summary>
    /// Provides the message texts for <see cref="TemplateErrorCode"/>s in the language of the user.
    /// English is built in; further languages are added with <see cref="AddLanguage"/> - either from a language
    /// package (e.g. <c>DocxTemplater.Localization.German</c>) or with your own dictionary. A dictionary does not
    /// have to be complete: a code without a text in the requested language falls back to the parent culture
    /// (<c>de-CH</c> → <c>de</c>) and finally to English. Adding a language twice merges the dictionaries, so
    /// single texts can be overridden.
    /// </summary>
    /// <remarks>
    /// The texts are composite format strings (<c>{0}</c>, <c>{1}</c>, ...); the arguments of each code are
    /// documented on <see cref="TemplateErrorCode"/>. Literal braces must be doubled (<c>{{</c>).
    /// </remarks>
    public partial class TemplateErrorMessages
    {
        private static readonly StringComparer CultureNameComparer = StringComparer.OrdinalIgnoreCase;

#if NET9_0_OR_GREATER
        private readonly System.Threading.Lock m_lock = new();
#else
        private readonly object m_lock = new();
#endif

        // copy-on-write: languages are usually registered once at startup but may be read concurrently
        private volatile Dictionary<string, IReadOnlyDictionary<TemplateErrorCode, string>> m_languages = new(CultureNameComparer);

        /// <summary>
        /// The instance used by <see cref="ProcessSettings"/> unless another one is configured. Languages added
        /// here are available for all documents.
        /// </summary>
        public static TemplateErrorMessages Default { get; } = new();

        /// <summary>
        /// Adds (or extends) the texts of a language. <paramref name="culture"/> can be neutral (<c>de</c>) or
        /// specific (<c>de-CH</c>); a specific culture falls back to its neutral culture for missing texts.
        /// </summary>
        /// <returns>This instance, to chain calls.</returns>
        public TemplateErrorMessages AddLanguage(CultureInfo culture, IReadOnlyDictionary<TemplateErrorCode, string> formats)
        {
            ArgumentNullException.ThrowIfNull(culture);
            ArgumentNullException.ThrowIfNull(formats);
            lock (m_lock)
            {
                var languages = new Dictionary<string, IReadOnlyDictionary<TemplateErrorCode, string>>(m_languages, CultureNameComparer);
                var merged = languages.TryGetValue(culture.Name, out var existing)
                    ? existing.ToDictionary(x => x.Key, x => x.Value)
                    : new Dictionary<TemplateErrorCode, string>();
                foreach (var format in formats)
                {
                    merged[format.Key] = format.Value;
                }
                languages[culture.Name] = merged;
                m_languages = languages;
            }
            return this;
        }

        /// <summary>
        /// The cultures that have been added with <see cref="AddLanguage"/>.
        /// </summary>
        public IReadOnlyCollection<string> Languages => m_languages.Keys.ToList();

        /// <summary>
        /// Returns the format string of <paramref name="code"/> for <paramref name="culture"/>, falling back to the
        /// parent cultures and finally to English. Returns <c>null</c> if the code is unknown.
        /// </summary>
        public virtual string GetFormat(TemplateErrorCode code, CultureInfo culture)
        {
            var languages = m_languages;
            for (var c = culture; c != null && !string.IsNullOrEmpty(c.Name); c = c.Parent)
            {
                if (languages.TryGetValue(c.Name, out var formats) && formats.TryGetValue(code, out var format))
                {
                    return format;
                }
            }
            return English.TryGetValue(code, out var english) ? english : null;
        }

        /// <summary>
        /// Formats the message of <paramref name="code"/> in the language of <paramref name="culture"/>.
        /// Arguments that are themselves errors (<see cref="OpenXmlTemplateException"/>, <see cref="TemplateSyntaxError"/>
        /// or a list of them) are rendered in the same language, so nested messages are translated as a whole.
        /// </summary>
        public string Format(TemplateErrorCode code, CultureInfo culture, params object[] arguments)
        {
            return Format(code, culture, (IReadOnlyList<object>)arguments);
        }

        /// <inheritdoc cref="Format(TemplateErrorCode, CultureInfo, object[])"/>
        public string Format(TemplateErrorCode code, CultureInfo culture, IReadOnlyList<object> arguments)
        {
            culture ??= CultureInfo.InvariantCulture;
            arguments ??= Array.Empty<object>();
            var args = new object[arguments.Count];
            for (int i = 0; i < args.Length; i++)
            {
                args[i] = RenderArgument(arguments[i], culture);
            }

            var format = GetFormat(code, culture);
            if (format != null && TryFormat(format, culture, args, out var message))
            {
                return message;
            }
            // a custom dictionary with a wrong number of placeholders must not break rendering - fall back to English
            if (English.TryGetValue(code, out var english) && TryFormat(english, culture, args, out message))
            {
                return message;
            }
            return args.Length == 0 ? code.ToString() : $"{code}: {string.Join(", ", args)}";
        }

        private static bool TryFormat(string format, CultureInfo culture, object[] args, out string message)
        {
            try
            {
                // the dictionaries use '\n' so they are platform independent; messages use the platform line break
                message = string.Format(culture, format.Replace("\n", Environment.NewLine), args);
                return true;
            }
            catch (FormatException)
            {
                message = null;
                return false;
            }
        }

        private object RenderArgument(object argument, CultureInfo culture)
        {
            return argument switch
            {
                OpenXmlTemplateException e => e.GetMessage(this, culture),
                TemplateSyntaxError error => error.ToString(this, culture),
                IEnumerable<TemplateSyntaxError> errors => string.Join(Environment.NewLine, errors.Select(x => x.ToString(this, culture))),
                Exception e => e.Message,
                _ => argument
            };
        }
    }
}
