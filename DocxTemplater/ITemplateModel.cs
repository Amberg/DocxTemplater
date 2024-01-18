namespace DocxTemplater
{
    public interface ITemplateModel
    {
        bool TryGetPropertyValue(string propertyName, out object value);
    }
}
