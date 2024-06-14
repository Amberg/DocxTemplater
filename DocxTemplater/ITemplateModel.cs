namespace DocxTemplater
{
    /// <summary>
    /// Interface for template model if a normal object or a dictionary is not suitable.
    /// </summary>
    public interface ITemplateModel
    {
        bool TryGetPropertyValue(string propertyName, out ValueWithMetadata value);
    }
}
