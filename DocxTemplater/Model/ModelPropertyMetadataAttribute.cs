using System;

namespace DocxTemplater.Model
{
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
    public class ModelPropertyAttribute : Attribute
    {
        public string DefaultFormatter { get; init; }

        public ModelPropertyAttribute()
        {
        }
    }
}
