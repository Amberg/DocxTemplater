using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocxTemplater.Charts
{
	public class ChartSeries
	{
		/// <summary>
		/// The name of the series.
		/// </summary>
		public string Name { get; set; }

		/// <summary>
		/// The values for the series.
		/// </summary>
		public List<double> Values { get; set; }

		/// <summary>
		/// Optional. If set, the series will be grouped under this category.
		/// Not required for all chart types.
		/// </summary>
		public List<string> CategoryNames { get; set; }
	}
}
