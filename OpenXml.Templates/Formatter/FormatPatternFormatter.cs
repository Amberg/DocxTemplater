using System;

namespace OpenXml.Templates.Formatter
{
    internal class FormatPatternFormatter : IFormatter
    {
        public bool CanHandle(Type type, string prefix)
        {
            if(prefix.ToUpper() == "FORMAT")
            {
               if(type == typeof(DateTime))
               {
                    return true;
               }
               if(type == typeof(decimal))
               {
                    return true;
               }
               if(type == typeof(int))
               {
                    return true;
               }
               if(type == typeof(float))
               {
                    return true;
               }
               if(type == typeof(double)) 
               {
                   return true;
               }
            }
            return type == typeof(DateTime) && prefix.ToUpper() == "FORMAT";
        }

        public string Format(object value, string prefix, string[] args)
        {
            if (args.Length != 1)
            {
                throw new OpenXmlTemplateException($"DateTime formatter requires exactly one argument, e.g. FORMAT(dd.MM.yyyy)");
            }
            if(value is DateTime dateTime)
            {
                return dateTime.ToString(args[0]);
            }
            if(value is decimal decimalValue)
            {
                return decimalValue.ToString(args[0]);
            }
            if(value is int intValue)
            {
                return intValue.ToString(args[0]);
            }
            if(value is float floatValue)
            {
                return floatValue.ToString(args[0]);
            }
            if(value is double doubleValue)
            {
                return doubleValue.ToString(args[0]);
            }
            return value.ToString();
        }
    }

}
