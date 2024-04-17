using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using ExcelExporter.Models;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace ExcelExporter.Helpers
{
    public static partial class Tools
    {
        public static void ApplyStyles(ExcelRange cell, SectionCellStyle sectionCell, bool isStart = false, bool isEnd = false)
        {
            Helpers.Tools.AssignByMembers(sectionCell?.Style.Numberformat ?? cell.Style.Numberformat, cell.Style.Numberformat, propertiesOnly: true);
            Helpers.Tools.AssignByMembers(sectionCell?.Style.Fill ?? cell.Style.Fill, cell.Style.Fill, propertiesOnly: true);
            Helpers.Tools.AssignByMembers(sectionCell?.Style.Font ?? cell.Style.Font, cell.Style.Font, propertiesOnly: true);

            ApplyBorderStyles(cell, sectionCell, isStart, isEnd);
            ApplyColors(cell, sectionCell);

            cell.Style.HorizontalAlignment = sectionCell?.Style.HorizontalAlignment ?? cell.Style.HorizontalAlignment;
            cell.Style.VerticalAlignment = sectionCell?.Style.VerticalAlignment ?? cell.Style.VerticalAlignment;
            cell.Style.WrapText = sectionCell?.Style.WrapText ?? cell.Style.WrapText;
        }

        public static void ApplyBorderStyles(ExcelRange cell, SectionCellStyle sectionCell, bool isStart, bool isEnd)
        {
            Helpers.Tools.AssignByMembers(sectionCell?.Style.Border.Left ?? cell.Style.Border.Left, cell.Style.Border.Left, propertiesOnly: true);
            if (!isStart)
            {
                Helpers.Tools.AssignByMembers(sectionCell?.Style.Border.Top ?? cell.Style.Border.Top, cell.Style.Border.Top, propertiesOnly: true);
            }
            Helpers.Tools.AssignByMembers(sectionCell?.Style.Border.Right ?? cell.Style.Border.Right, cell.Style.Border.Right, propertiesOnly: true);
            if (!isEnd)
            {
                Helpers.Tools.AssignByMembers(sectionCell?.Style.Border.Bottom ?? cell.Style.Border.Bottom, cell.Style.Border.Bottom, propertiesOnly: true);
            }
        }

        public static void ApplyColors(ExcelRange cell, SectionCellStyle sectionCell)
        {
            if (cell.Style.Fill.PatternType != ExcelFillStyle.None)
            {
                if (TryGetColor(sectionCell?.Style.Fill.PatternColor.Rgb ?? cell.Style.Fill.PatternColor.Rgb, out Color patternColor))
                {
                    cell.Style.Fill.PatternColor.SetColor(patternColor);
                }
                if (TryGetColor(sectionCell?.Style.Fill.BackgroundColor.Rgb ?? cell.Style.Fill.BackgroundColor.Rgb, out Color backgroundColor))
                {
                    cell.Style.Fill.BackgroundColor.SetColor(backgroundColor);
                }
            }

            if (TryGetColor(sectionCell?.Style.Font.Color.Rgb ?? cell.Style.Font.Color.Rgb, out Color fontColor))
            {
                cell.Style.Font.Color.SetColor(fontColor);
            }
        }

        public static bool TryGetColor(string hexValue, out Color color)
        {
            if (int.TryParse(hexValue, System.Globalization.NumberStyles.HexNumber, null, out int argb))
            {
                color = Color.FromArgb(argb);
                return true;
            }
            color = default;
            return false;
        }

        public static bool IsRowEmpty(ExcelWorksheet worksheet, int row)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (!string.IsNullOrWhiteSpace(worksheet.Cells[row, col].Text))
                {
                    return false;
                }
            }
            return true;
        }

        public static void IdentifyMergedRanges(ExcelWorksheet worksheet, Section section)
        {
            foreach (var mergeAddress in worksheet.MergedCells)
            {
                var mergedRange = worksheet.Cells[mergeAddress];
                if (IsRangeWithinSection(section, mergedRange))
                {
                    section.MergedCells.Add(new ExcelAddress(mergeAddress));
                }
            }
        }

        public static bool IsRangeWithinSection(Section section, ExcelRange mergedRange)
        {
            return mergedRange.Start.Row >= section.StartRow + 1 &&
                   mergedRange.End.Row <= section.EndRow - 1 &&
                   mergedRange.Start.Column >= section.StartCol &&
                   mergedRange.End.Column <= section.EndCol;
        }
    }
}
