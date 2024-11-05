using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelExporter.Models
{
    public class Route
    {
        public Route() { }

        public Route(string defName, int depth, int index)
        {
            DefName = defName;
            Depth = depth;
            Index = index;
        }

        public string DefName { get; set; }
        public int Depth { get; set; }
        public int? Index { get; set; }
    }
}
