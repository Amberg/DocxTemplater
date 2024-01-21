using System.Collections.Generic;

namespace DocxTemplater
{
    public interface IDynamicTable
    {
        public IEnumerable<object> Headers { get; }


        public IEnumerable<IEnumerable<object>> Rows { get; }

    }
}
