using System.Collections.Generic;

namespace DocxTemplater
{
    public interface IModelLookup
    {
        IReadOnlyDictionary<string, object> Models { get; }
        void Add(string prefix, object model);
        IVariableScope OpenScope();
        object GetValue(string variableName);
        object GetScopeParentLevel(int parentLevel);
    }
}