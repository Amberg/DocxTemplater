using System.Globalization;

namespace DocxTemplater
{
    public class ProcessSettings
    {

        public CultureInfo Culture
        {
            get;
            set;
        } = CultureInfo.CurrentUICulture;

        public BindingErrorHandling BindingErrorHandling
        {
            get;
            set;
        } = BindingErrorHandling.ThrowException;

        public static ProcessSettings Default { get; } = new ProcessSettings() { Culture = null }; // will use current ui culture
    }
}
