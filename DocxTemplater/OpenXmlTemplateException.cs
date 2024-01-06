using System;

namespace DocxTemplater
{
    [Serializable]
    public class OpenXmlTemplateException : Exception
    {
        public OpenXmlTemplateException(string message) : base(message) { }
        public OpenXmlTemplateException(string message, System.Exception inner) : base(message, inner) { }
    }
}
