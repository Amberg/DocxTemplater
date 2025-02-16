using System.Collections.Generic;

namespace DocxTemplater.Extensions.Charts
{
    public class ChartData
    {
        /// <summary>
        /// Empty to remove the chart title.
        /// </summary>
        public string ChartTitle { get; set; }

        public List<ChartSeries> Series { get; set; }

        /// <summary>
        /// X - Labels for the series.
        /// </summary>
        public IEnumerable<string> Categories { get; set; }
    }

}