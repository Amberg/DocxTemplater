using System.Collections.Generic;

namespace DocxTemplater.Extensions.Charts
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
        public IEnumerable<double> Values { get; set; }
    }
}