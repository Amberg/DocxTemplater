namespace DocxTemplater
{
    /// <summary>
    /// Rewrites the content of a recognized tag before it is parsed - e.g. to decode what an editor encoded
    /// (<c>&amp;gt;</c>) or to replace typographic quotes with straight ones. Registered with
    /// <c>RegisterTagContentPreprocessor</c> of the <see cref="ProcessSettings"/> or of the settings of a text template.
    /// </summary>
    /// <remarks>
    /// Only the parts listed in <see cref="TagContentKind"/> are passed; text outside of tags, variable names and tags
    /// the pattern matcher does not recognize (<see cref="TemplateErrorCode.InvalidTag"/>) never are. Error positions
    /// refer to the template as written.
    /// </remarks>
    public interface ITagContentPreprocessor
    {
        /// <returns>The content to parse instead of <paramref name="content"/>; <paramref name="content"/> to keep it.</returns>
        string Process(string content, TagContentKind kind);
    }
}
