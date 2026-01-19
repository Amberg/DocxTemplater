using System;

namespace DocxTemplater.Model
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ModelPropertyAttribute : Attribute
    {
        public string DefaultFormatter { get; set; }

        public ModelPropertyAttribute()
        {
        }
    }
}
