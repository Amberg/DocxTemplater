using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater
{
    public class DynamicTable : IDynamicTable
    {
        private readonly List<Dictionary<object, object>> m_rows;
        public IEnumerable<object> Headers => m_rows.SelectMany(x => x.Keys).Distinct().ToList();

        public IEnumerable<IEnumerable<object>> Rows => m_rows.Select(x => x.Values.ToList()).ToList();

        public DynamicTable()
        {
            m_rows = new List<Dictionary<object, object>>();
        }

        public void AddRow(Dictionary<object, object> row)
        {
            m_rows.Add(row);
        }
    }
}
