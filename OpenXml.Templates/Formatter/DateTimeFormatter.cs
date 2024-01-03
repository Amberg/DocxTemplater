using System;
using System.Collections.Generic;
using System.Text;

namespace OpenXml.Templates.Formatter
{
    internal class DateTimeFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            return type == typeof(DateTime) && prefix.ToUpper() == "FORMAT";
        }

        public string Format(object value, string prefix, string[] args)
        {
            if (args.Length != 1)
            {
                throw new OpenXmlTemplateException($"DateTime formatter requires exactly one argument, e.g. FORMAT(dd.MM.yyyy)");
            }
            return ((DateTime)value).ToString(args[0]);
        }
    }

}
