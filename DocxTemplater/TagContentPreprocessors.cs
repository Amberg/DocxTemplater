using System;

namespace DocxTemplater
{
    /// <summary>The <see cref="ITagContentPreprocessor"/>s of the settings, applied in registration order. Immutable.</summary>
    internal sealed class TagContentPreprocessors
    {
        private readonly ITagContentPreprocessor[] m_preprocessors;

        private TagContentPreprocessors(ITagContentPreprocessor[] preprocessors)
        {
            m_preprocessors = preprocessors;
        }

        public static TagContentPreprocessors None { get; } = new([]);

        public bool IsEmpty => m_preprocessors.Length == 0;

        /// <summary>The preprocessors of the settings of a Word template context.</summary>
        public static TagContentPreprocessors Of(ITemplateProcessingContext context)
        {
            return context.ProcessSettings?.TagContentPreprocessors ?? None;
        }

        public TagContentPreprocessors With(ITagContentPreprocessor preprocessor)
        {
            return new TagContentPreprocessors([.. m_preprocessors, preprocessor]);
        }

        public string Process(string content, TagContentKind kind)
        {
            foreach (var preprocessor in m_preprocessors)
            {
                content = preprocessor.Process(content, kind)
                          ?? throw new InvalidOperationException($"{preprocessor.GetType().Name}.{nameof(ITagContentPreprocessor.Process)} returned null.");
            }
            return content;
        }
    }
}
