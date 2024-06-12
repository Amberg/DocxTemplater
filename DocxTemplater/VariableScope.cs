using System;

namespace DocxTemplater
{
    public interface IVariableScope : IDisposable
    {
        void AddVariable(string name, object value);
    }
}
