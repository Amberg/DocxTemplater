using System.Globalization;

namespace DocxTemplater
{
    public class ProcessSettings
    {
        private string _openDelimiter = "{{";
        private string _closeDelimiter = "}}";
        private PatternMatcher _patternMatcher;

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
        /// The opening delimiter for template tags (default: "{{").
        /// Must be at least 2 characters and different from <see cref="CloseDelimiter"/>.
        /// Example: set to "&lt;&lt;" to use &lt;&lt;variable&gt;&gt; syntax.
        /// </summary>
        public string OpenDelimiter
        {
            get => _openDelimiter;
            set
            {
                _openDelimiter = value;
                _patternMatcher = null;
            }
        }

        /// <summary>
        /// The closing delimiter for template tags (default: "}}").
        /// Must be at least 2 characters and different from <see cref="OpenDelimiter"/>.
        /// Example: set to "&gt;&gt;" to use &lt;&lt;variable&gt;&gt; syntax.
        /// </summary>
        public string CloseDelimiter
        {
            get => _closeDelimiter;
            set
            {
                _closeDelimiter = value;
                _patternMatcher = null;
            }
        }

        internal PatternMatcher PatternMatcher
            => _patternMatcher ??= new PatternMatcher(_openDelimiter, _closeDelimiter);

        public static ProcessSettings Default => new();
    }
}
