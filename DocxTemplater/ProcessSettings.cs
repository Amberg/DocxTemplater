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

        /// <summary>
        /// When enabled, this option removes leading or trailing newlines around template directives (e.g., {{#...}}, {{/}}) 
        /// from the final output. This allows templates to be more readable without affecting rendered formatting.
        /// default: false
        /// </summary>
        public bool IgnoreLineBreaksAroundTags { get; set; }

        /// <summary>
        /// When enabled, this option removes paragraphs that only contain conditional blocks or loops when these blocks
        /// don't render any content (e.g., when a condition is false or a collection is empty).
        /// This helps to avoid empty paragraphs in the final document.
        /// default: false
        /// </summary>
        public bool RemoveParagraphsContainingOnlyBlocks { get; set; }

        public static ProcessSettings Default { get; } = new();
    }
}
