using System;

namespace DocxTemplater
{
    public interface IScriptCompiler
    {
        ProcessSettings ProcessSettings { get; }
        Func<bool> CompileScript(string scriptAsString);
        Func<object> CompileExpression(string scriptAsString);
    }
}