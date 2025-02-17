using System;

namespace DocxTemplater.Model
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ModelPropertyAttribute : Attribute
    {
        public string DefaultFormatter { get; init; }

        public ModelPropertyAttribute()
        {
        }
    }
}
