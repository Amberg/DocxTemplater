using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

        public void ApplyFormat(string modelPath, object value, string prefix, string[] args, Text target)
        {
            var text = value.ToString();
            if (args.Length != 1)
            {
                throw new OpenXmlTemplateException($"DateTime formatter requires exactly one argument, e.g. FORMAT(dd.MM.yyyy)");
            }
            
            if (value is DateTime dateTime)
            {
                text = dateTime.ToString(args[0]);
            }
            else if (value is decimal decimalValue)
            {
                text = decimalValue.ToString(args[0]);
            }
            else if (value is int intValue)
            {
                text = intValue.ToString(args[0]);
            }
            else if (value is float floatValue)
            {
                text = floatValue.ToString(args[0]);
            }
            else if (value is double doubleValue)
            {
                text = doubleValue.ToString(args[0]);
            } 
            target.Text = text;
        }
    }

}
