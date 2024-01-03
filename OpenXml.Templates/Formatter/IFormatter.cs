using System;

namespace OpenXml.Templates.Formatter
{
    internal interface IFormatter
    {
        public bool CanHandle(Type type, string prefix);

        public string Format(object value, string prefix, string[] args);
    }
}
