using System.Globalization;

namespace DocxTemplater
{
    public class ProcessSettings
    {

        /// <summary>
        /// Output culture of the document
        /// </summary>
        public CultureInfo Culture { get; set; } = CultureInfo.CurrentUICulture;

        public BindingErrorHandling BindingErrorHandling { get; set; } = BindingErrorHandling.ThrowException;

        public static ProcessSettings Default { get; } = new();
    }
}
