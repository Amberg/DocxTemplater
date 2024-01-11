using System.Globalization;

namespace DocxTemplater
{
    public class ProcessSettings
    {
        private CultureInfo m_culture;

        public CultureInfo Culture
        {
            get => m_culture ?? CultureInfo.CurrentUICulture;
            set => m_culture = value;
        }

        public static ProcessSettings Default { get; } = new ProcessSettings() { Culture = null }; // will use current ui culture
    }
}
