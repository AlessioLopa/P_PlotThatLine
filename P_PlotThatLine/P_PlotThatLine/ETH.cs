using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace P_PlotThatLine
{
    internal class ETH
    {
        public ETH(string start, string end, string open, string high, string low, string close)
        {
            this.start = Convert.ToDateTime(start);
            this.end = Convert.ToDateTime(end);
            this.open = Convert.ToSingle(open);
            this.high = Convert.ToSingle(high);
            this.low = Convert.ToSingle(low);
            this.close = Convert.ToSingle(close);
        }

        public DateTime start { get; set; }
        public DateTime end { get; set; }
        public double open { get; set; }
        public double high { get; set; }
        public double low { get; set; }
        public double close { get; set; }

    }
}
