using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;

namespace DocxTemplater
{
    /// <summary>
    /// Provides the message texts for <see cref="TemplateErrorCode"/>s in the language of the user.
    /// English is built in. Other languages come from language packages (<c>DocxTemplater.Localization.de</c>,
    /// <c>DocxTemplater.Localization.de-CH</c>, <c>DocxTemplater.Localization.fr</c>, ...) that are loaded
    /// automatically when installed - see <see cref="AutoLoadLanguagePackages"/> - or are added with
    /// <see cref="AddLanguage(CultureInfo, IReadOnlyDictionary{TemplateErrorCode, string})"/> from your own dictionary.
    /// A dictionary does not have to be complete: a code without a text in the requested language falls back to the
    /// parent culture (<c>de-CH</c> → <c>de</c>) and finally to English. Adding a language twice merges the
    /// dictionaries, so single texts can be overridden.
    /// </summary>
    /// <remarks>
    /// The texts are composite format strings (<c>{0}</c>, <c>{1}</c>, ...); the arguments of each code are
    /// documented on <see cref="TemplateErrorCode"/>. Literal braces must be doubled (<c>{{</c>).
    /// </remarks>
    public partial class TemplateErrorMessages
    {
        /// <summary>
        /// Prefix of the assembly name of a language package; the culture name is appended.
        /// </summary>
        public const string LanguagePackageAssemblyPrefix = "DocxTemplater.Localization.";

        private static readonly StringComparer CultureNameComparer = StringComparer.OrdinalIgnoreCase;

#if NET9_0_OR_GREATER
        private readonly System.Threading.Lock m_lock = new();
#else
        private readonly object m_lock = new();
#endif

        // copy-on-write: languages are usually registered once at startup but may be read concurrently
        private volatile Dictionary<string, IReadOnlyDictionary<TemplateErrorCode, string>> m_languages = new(CultureNameComparer);

        // cultures for which a language package has been looked up (successfully or not)
        private volatile HashSet<string> m_probedCultures = new(CultureNameComparer);

        /// <summary>
        /// The instance used by <see cref="ProcessSettings"/> unless another one is configured. Languages added
        /// here are available for all documents.
        /// </summary>
        public static TemplateErrorMessages Default { get; } = new();

        /// <summary>
        /// Whether an installed language package (an assembly named <c>DocxTemplater.Localization.&lt;culture&gt;</c>
        /// containing an <see cref="ITemplateLanguagePack"/>) is loaded automatically the first time a message is
        /// formatted for its culture. The lookup happens once per culture name and the result is cached.
        /// Disable it in trimmed applications and register the packs with <see cref="AddLanguage(ITemplateLanguagePack)"/>.
        /// default: true
        /// </summary>
        public bool AutoLoadLanguagePackages { get; set; } = true;

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
        /// Adds (or extends) the texts of a language package, e.g. <c>AddLanguage(new GermanErrorMessages())</c>.
        /// Only needed when <see cref="AutoLoadLanguagePackages"/> is disabled.
        /// </summary>
        /// <returns>This instance, to chain calls.</returns>
        public TemplateErrorMessages AddLanguage(ITemplateLanguagePack pack)
        {
            ArgumentNullException.ThrowIfNull(pack);
            return AddLanguage(pack.Culture, pack.Formats);
        }

        /// <summary>
        /// The cultures that have been added with <see cref="AddLanguage(CultureInfo, IReadOnlyDictionary{TemplateErrorCode, string})"/>
        /// or loaded from a language package.
        /// </summary>
        public IReadOnlyCollection<string> Languages => m_languages.Keys.ToList();

        /// <summary>
        /// Returns the format string of <paramref name="code"/> for <paramref name="culture"/>, falling back to the
        /// parent cultures and finally to English. Returns <c>null</c> if the code is unknown.
        /// </summary>
        public virtual string GetFormat(TemplateErrorCode code, CultureInfo culture)
        {
            for (var c = culture; c != null && !string.IsNullOrEmpty(c.Name); c = c.Parent)
            {
                if (!m_languages.ContainsKey(c.Name) && AutoLoadLanguagePackages)
                {
                    TryLoadLanguagePackage(c);
                }
                if (m_languages.TryGetValue(c.Name, out var formats) && formats.TryGetValue(code, out var format))
                {
                    return format;
                }
            }
            // texts added for the invariant culture override the built-in English
            if (m_languages.TryGetValue(CultureInfo.InvariantCulture.Name, out var invariant) && invariant.TryGetValue(code, out var overridden))
            {
                return overridden;
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

        /// <summary>
        /// Loads the language package of <paramref name="culture"/> by convention: the assembly
        /// <c>DocxTemplater.Localization.&lt;culture name&gt;</c> deployed with the application. Each culture is
        /// probed once; a missing package is remembered so the lookup is not repeated.
        /// </summary>
        private void TryLoadLanguagePackage(CultureInfo culture)
        {
            if (m_probedCultures.Contains(culture.Name))
            {
                return;
            }
            // load inside the lock, so a concurrent caller waits for the texts instead of seeing "probed, nothing found"
            lock (m_lock)
            {
                if (m_probedCultures.Contains(culture.Name))
                {
                    return;
                }
                foreach (var pack in LoadLanguagePacks(culture.Name))
                {
                    AddLanguage(pack);
                }
                m_probedCultures = new HashSet<string>(m_probedCultures, CultureNameComparer) { culture.Name };
            }
        }

        /// <summary>
        /// A broken or incompatible package must not turn an error message into a crash of the rendering, so
        /// everything that can go wrong while loading or instantiating a pack is treated as "no package".
        /// </summary>
        private static IReadOnlyList<ITemplateLanguagePack> LoadLanguagePacks(string cultureName)
        {
            try
            {
                var assembly = Assembly.Load(new AssemblyName(LanguagePackageAssemblyPrefix + cultureName));
                return assembly.GetExportedTypes()
                    .Where(t => !t.IsAbstract && typeof(ITemplateLanguagePack).IsAssignableFrom(t) && t.GetConstructor(Type.EmptyTypes) != null)
                    .Select(t => (ITemplateLanguagePack)Activator.CreateInstance(t))
                    .ToList();
            }
            catch (Exception e) when (e is FileNotFoundException or FileLoadException or BadImageFormatException
                                          or ReflectionTypeLoadException or TypeLoadException or TypeInitializationException
                                          or TargetInvocationException or MemberAccessException or NotSupportedException)
            {
                return Array.Empty<ITemplateLanguagePack>();
            }
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
