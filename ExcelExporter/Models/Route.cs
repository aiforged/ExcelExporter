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

        public Route(int depth, int index)
        {
            Depth = depth;
            Index = index;
        }

        public int Depth { get; set; }
        public int? Index { get; set; }
    }
}
