using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExporter.Models
{
    public class Section
    {
        public Section()
        {
        }

        public Section(string name, int startRow)
        {
            Name = name;
            StartRow = startRow;
        }

        public Section(string name, string parentName, int startRow, int startCol, int endCol)
        {
            Name = name;
            ParentName = parentName;
            StartRow = startRow;
            StartCol = startCol;
            EndCol = endCol;
        }

        public Section(string name, string parentName, int startRow, int startCol, int endRow, int endCol, object cells, List<SectionCellStyle> cellStyles, List<ExcelAddress> mergedCells)
        {
            Name = name;
            ParentName = parentName;
            StartRow = startRow;
            StartCol = startCol;
            EndRow = endRow;
            EndCol = endCol;
            Cells = cells;
            CellStyles = cellStyles;
            MergedCells = mergedCells;
        }

        public string Name { get; set; }
        public string ParentName { get; set; }
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int EndRow { get; set; }
        public int EndCol { get; set; }
        public object Cells { get; set; }
        public List<SectionCellStyle> CellStyles { get; set; } = new List<SectionCellStyle>();
        public List<ExcelAddress> MergedCells { get; set; } = new List<ExcelAddress>();

        public ExcelAddress Range => new ExcelAddress(StartRow, StartCol, EndRow, EndCol);

        public bool IsRangeWithin(string innerRange)
        {
            return IsRangeWithin(new ExcelAddress(innerRange));
        }

        public bool IsRangeWithin(ExcelAddress innerRange)
        {
            return innerRange.Start.Row >= Range.Start.Row &&
                   innerRange.End.Row <= Range.End.Row &&
                   innerRange.Start.Column >= Range.Start.Column &&
                   innerRange.End.Column <= Range.End.Column;
        }
    }

    public class SectionCellStyle
    {
        public SectionCellStyle(int row, int col, ExcelStyle style)
        {
            Row = row;
            Col = col;
            Style = style;
        }

        public int Row { get; set; }
        public int Col { get; set; }
        public ExcelStyle Style { get; set; }
    }
}
