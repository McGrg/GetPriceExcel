using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetPriceExcel
{
    class GridRows
    {
        public string Works { get; set; }
        public string Volume { get; set; }
        public string Materials { get; set; }
        public string Smr { get; set; }
        public GridRows (string Works, string Volume, string Materials, string Smr)
        {
            this.Works = Works;
            this.Volume = Volume;
            this.Materials = Materials;
            this.Smr = Smr;
        }
    }
}
