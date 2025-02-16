using DocxTemplater.Model;
using System.Collections.Generic;
using System.Linq;

namespace DocxTemplater
{
    public class DynamicTable : IDynamicTable
    {
        private readonly IEqualityComparer<object> m_headerComparer;
        private readonly List<Dictionary<object, object>> m_rows;
        public IEnumerable<object> Headers => m_rows.SelectMany(x => x.Keys).Distinct(m_headerComparer).ToList();

        public IEnumerable<IEnumerable<object>> Rows => m_rows.Select(x => x.Values.ToList()).ToList();

        public DynamicTable(IEqualityComparer<object> headerComparer = null)
        {
            m_headerComparer = headerComparer ?? EqualityComparer<object>.Default;
            m_rows = new List<Dictionary<object, object>>();
        }

        public void AddRow(Dictionary<object, object> row)
        {
            m_rows.Add(row);
        }
    }
}
