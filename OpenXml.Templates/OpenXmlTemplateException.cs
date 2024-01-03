using System;

namespace OpenXml.Templates
{
    [Serializable]
    public class OpenXmlTemplateException : Exception
    {
        public OpenXmlTemplateException(string message) : base(message) { }
        public OpenXmlTemplateException(string message, System.Exception inner) : base(message, inner) { }
    }
}
