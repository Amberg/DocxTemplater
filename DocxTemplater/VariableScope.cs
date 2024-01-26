using System;

namespace DocxTemplater
{
    internal interface IVariableScope : IDisposable
    {
        void AddVariable(string name, object value);
    }
}
