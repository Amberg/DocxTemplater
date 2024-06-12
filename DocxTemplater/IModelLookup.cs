using System.Collections.Generic;

namespace DocxTemplater
{
    public record struct ValueMetadata(string DefaultFormatter);

    public record struct ValueWithMetadata(object Value, ValueMetadata Metadata);

    public interface IModelLookup
    {
        IReadOnlyDictionary<string, object> Models { get; }
        void Add(string prefix, object model);
        IVariableScope OpenScope();

        object GetValue(string variableName);

        ValueWithMetadata GetValueWithMetadata(string variableName);

        object GetScopeParentLevel(int parentLevel);
    }
}