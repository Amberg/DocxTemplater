using System.Collections.Generic;

namespace DocxTemplater.Model
{
    public interface IDynamicTable
    {
        IEnumerable<object> Headers { get; }

        IEnumerable<IEnumerable<object>> Rows { get; }
    }
}
