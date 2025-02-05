using System.Collections.Generic;

namespace DocxTemplater.Charts
{
	public class ChartData
	{

		/// <summary>
		/// Empty to remove the chart title.
		/// </summary>
		public string ChartTitle { get; set; }

		public List<ChartSeries> Series { get; set; }
	}

}
