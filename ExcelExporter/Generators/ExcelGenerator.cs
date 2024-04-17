using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using AIForged.API;
using AIForged.DAL.Models;

using ExcelExporter.Models;

using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Ranges;
using OfficeOpenXml.Style;

namespace ExcelExporter.Generators
{
    public class ExcelGenerator
    {
        private ILogger _logger;
        private int _iteration = 0;
        private Stack<Section> _templateStack = new Stack<Section>();

        public ExcelGenerator(ILogger logger) 
        { 
            this._logger = logger;
        }

        public void GenerateExcelFile(string templatePath, string outputPath, ICollection<DocumentParameterViewModel> data)
        {
            FileInfo templateFile = new FileInfo(templatePath);
            FileInfo newFile = new FileInfo(outputPath);
            _iteration = 0;

            using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Template"];
                int currentRow = 1;
                int templateRows = ParseTemplate(worksheet, currentRow);

                foreach (var child in data)
                {
                    currentRow = ProcessSections(worksheet, child, currentRow);

                    if (data.ToList().IndexOf(child) == data.Count - 1) break;

                    CloneTemplate(worksheet, currentRow);
                    _iteration++;
                }

                SavePackage(package);
            }
        }

        private int ParseTemplate(ExcelWorksheet worksheet, int startRow)
        {
            int finalEndRow = startRow;
            _templateStack = new Stack<Section>();

            for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text;
                    var match = Regex.Match(cellValue, @"\{([a-zA-Z\:\|\.]+)\s(start|end)\}");

                    if (match.Success)
                    {
                        HandleSectionMatch(worksheet, row, match);
                        finalEndRow = row;
                        col = worksheet.Dimension.End.Column; // Force next row
                    }
                }
            }
            return finalEndRow;
        }

        private void HandleSectionMatch(ExcelWorksheet worksheet, int row, Match match)
        {
            string sectionName = match.Groups[1].Value;
            string sectionType = match.Groups[2].Value;

            if (sectionType == "start")
            {
                AddSectionStart(worksheet, row, sectionName);
            }
            else if (sectionType == "end" && _templateStack.Count > 0)
            {
                CompleteSectionEnd(worksheet, row, sectionName);
            }
        }

        private void AddSectionStart(ExcelWorksheet worksheet, int row, string sectionName)
        {
            int startColumn = FindStartColumn(worksheet, row);
            int endColumn = FindEndColumn(worksheet, row);

            _templateStack.Push(new Section(sectionName, row, startColumn, endColumn));
        }

        private int FindStartColumn(ExcelWorksheet worksheet, int row)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (worksheet.Cells[row, col].Style.Border.Left.Style == ExcelBorderStyle.Thick ||
                    worksheet.Cells[row, col].Style.Border.Left.Style == ExcelBorderStyle.Medium)
                {
                    return col;
                }
            }
            return 1;
        }

        private int FindEndColumn(ExcelWorksheet worksheet, int row)
        {
            for (int col = worksheet.Dimension.End.Column; col >= 1; col--)
            {
                if (worksheet.Cells[row, col].Style.Border.Right.Style == ExcelBorderStyle.Thick ||
                    worksheet.Cells[row, col].Style.Border.Right.Style == ExcelBorderStyle.Medium)
                {
                    return col;
                }
            }
            return worksheet.Dimension.End.Column;
        }

        private void CompleteSectionEnd(ExcelWorksheet worksheet, int row, string sectionName)
        {
            var section = _templateStack.FirstOrDefault(s => s.Name.Equals(sectionName));

            if (section != null)
            {
                section.EndRow = row;
                section.Cells = worksheet.Cells[section.StartRow, 1, row, worksheet.Dimension.End.Column].Value;
                CaptureSectionStyles(worksheet, section);
                Helpers.Tools.IdentifyMergedRanges(worksheet, section);
            }
        }

        private void CaptureSectionStyles(ExcelWorksheet worksheet, Section section)
        {
            foreach (var cell in worksheet.Cells[section.StartRow, 1, section.EndRow, worksheet.Dimension.End.Column])
            {
                section.CellStyles.Add(new SectionCellStyle(cell.Start.Row, cell.Start.Column, cell.Style));
            }
        }

        private int ProcessSections(ExcelWorksheet worksheet, DocumentParameterViewModel data, int startRow)
        {
            int rowExpansion = 0;

            var sections = Helpers.Tools.CloneStack(_templateStack);
            Section previousSection = null;
            int maxEndRow = 0;

            while (sections.Count > 0)
            {
                var section = sections.Pop();
                (int currentStartRow, int currentEndRow, rowExpansion) = ProcessSection(worksheet, data, startRow, rowExpansion, section, previousSection);
                previousSection = section;

                if (currentEndRow > maxEndRow)
                {
                    maxEndRow = currentEndRow;
                }
            }

            RemoveEmptyRows(worksheet);

            return worksheet.Dimension.End.Row + 1;
        }

        private (int currentStartRow, int currentEndRow, int rowExpansion) ProcessSection(ExcelWorksheet worksheet, DocumentParameterViewModel data, int startRow, int rowExpansion, Section section, Section previousSection)
        {
            if (previousSection == null)
            {
                return PopulateSection(worksheet, section, data, section.StartRow + (startRow - 1), section.EndRow + (startRow - 1) + rowExpansion);
            }
            else if (IsNestedSection(previousSection, section, rowExpansion))
            {
                return PopulateSection(worksheet, section, data, section.StartRow + (startRow - 1), section.EndRow + (startRow - 1) + rowExpansion);
            }
            else if (IsAppearingBefore(previousSection, section))
            {
                return PopulateSection(worksheet, section, data, section.StartRow + (startRow - 1), section.EndRow + (startRow - 1));
            }
            else
            {
                return PopulateSection(worksheet, section, data, section.StartRow + (startRow - 1) + rowExpansion, section.EndRow + (startRow - 1));
            }
        }

        private bool IsNestedSection(Section previousSection, Section section, int rowExpansion)
        {
            return previousSection.EndRow + rowExpansion < section.EndRow + rowExpansion && previousSection.StartRow > section.StartRow;
        }

        private bool IsAppearingBefore(Section previousSection, Section section)
        {
            return previousSection.StartRow > section.EndRow;
        }

        private void RemoveEmptyRows(ExcelWorksheet worksheet)
        {
            for (int row = worksheet.Dimension.End.Row; row >= 1; row--)
            {
                if (Helpers.Tools.IsRowEmpty(worksheet, row))
                {
                    worksheet.DeleteRow(row);
                }
                else
                {
                    break;
                }
            }
        }

        private void CloneTemplate(ExcelWorksheet worksheet, int initialRow)
        {
            var sections = Helpers.Tools.CloneStack(_templateStack);

            while (sections.Count > 0)
            {
                var section = sections.Pop();
                if (!IsSectionNested(sections, section))
                {
                    DuplicateSection(worksheet, initialRow, section);
                }
            }

            MergeClonedSections(worksheet, initialRow);
        }

        private bool IsSectionNested(Stack<Section> sections, Section section)
        {
            return sections.Any(s => s != section && s.IsRangeWithin(section.Range));
        }

        private void DuplicateSection(ExcelWorksheet worksheet, int initialRow, Section section)
        {
            int startRow = section.StartRow + initialRow;
            int endRow = section.EndRow + initialRow;
            int sectionRowCount = endRow - startRow;

            worksheet.InsertRow(startRow, sectionRowCount);
            worksheet.Cells[startRow, 1, endRow, section.EndCol].Value = section.Cells;

            ApplySectionStyles(worksheet, section, startRow);
        }

        private void ApplySectionStyles(ExcelWorksheet worksheet, Section section, int startRow)
        {
            for (int i = section.StartRow; i <= section.EndRow; i++)
            {
                for (int newCol = section.StartCol; newCol <= section.EndCol; newCol++)
                {
                    var sectionCell = section.CellStyles.FirstOrDefault(c => c.Row == i && c.Col == newCol);
                    var cell = worksheet.Cells[startRow + (i - section.StartRow), newCol, startRow + (i - section.StartRow), newCol];

                    Helpers.Tools.ApplyStyles(cell, sectionCell, i == startRow, i == section.EndRow);
                }
            }
        }

        private void MergeClonedSections(ExcelWorksheet worksheet, int initialRow)
        {
            var sections = Helpers.Tools.CloneStack(_templateStack);

            while (sections.Count > 0)
            {
                var section = sections.Pop();
                foreach (var mergeAddress in section.MergedCells)
                {
                    var newMergeAddress = new ExcelAddress(mergeAddress.Start.Row + initialRow,
                        mergeAddress.Start.Column,
                        mergeAddress.End.Row + initialRow,
                        mergeAddress.End.Column);

                    try
                    {
                        worksheet.Cells[newMergeAddress.Address].Merge = true;
                    }
                    catch
                    {
                        // Handle exceptions if needed
                    }
                }
            }
        }

        private (int startRow, int endRow, int rowExpansion) PopulateSection(ExcelWorksheet worksheet, Section sectionObject, DocumentParameterViewModel data, int startRow, int endRow)
        {
            int originalStartRow = startRow;
            int originalSectionRowCount = endRow - startRow;
            int sectionRowCount = endRow - startRow;
            //worksheet.DeleteRow(endRow, 1);
            //worksheet.DeleteRow(startRow, 1);

            if (sectionObject.Name.Contains(":") || sectionObject.Name.Contains("|"))
            {
                var paths = sectionObject.Name.Split("|", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

                (startRow, endRow) = PopulateSectionDataRoute(worksheet, sectionObject, data, startRow, endRow, paths.FirstOrDefault(), parentSection: null, routes: null, index: null, expand: false, paths.Where(r => r != paths.FirstOrDefault()).ToArray());
            }
            else
            {
                List<DocumentParameterViewModel> sectionData = Helpers.Tools.GetParameters(sectionObject.Name, data);

                if (sectionData.Count == 0)
                {
                    sectionData.Add(data);
                }

                foreach (DocumentParameterViewModel section in sectionData)
                {
                    (startRow, endRow) = PopulateSectionData(worksheet, sectionObject, section, startRow, endRow);
                }
            }

            return (startRow, endRow, endRow - originalStartRow - originalSectionRowCount);
        }

        private (int startRow, int endRow) PopulateSectionDataRoute(ExcelWorksheet worksheet, Section sectionObject, DocumentParameterViewModel data, int startRow, int endRow, string route, DocumentParameterViewModel parentSection = null, List<Route> routes = null, int? index = null, bool expand = false, params string[] path)
        {
            var section = Helpers.Tools.GetParameterRoute(route, data, index: index, checkColumnChildren: true, routes: routes?.ToArray());

            if (section is null) return (startRow, endRow);
            var rowIndexes = section.Children
                        .SelectMany(c => c.Children.Select(cc => cc.RowIndex ?? cc.Index))
                        .Distinct()
                        .OrderBy(c => c)
                        .ToList();

            for (int i = 0; i <= rowIndexes.Max().Value; i++)
            {
                routes ??= new List<Route>();
                var depth = route.Split(":").Count() - 1;
                var foundRoute = routes.FirstOrDefault(r => r.Depth == depth);

                if (foundRoute is null)
                {
                    routes.Add(new Route()
                    {
                        Depth = depth,
                        Index = index,
                    });
                }
                else
                {
                    foundRoute.Index = index;
                }

                if (path is null || path.Length == 0)
                {
                    (startRow, endRow) = PopulateSectionData(worksheet,
                        sectionObject,
                        parentSection ?? data,
                        startRow,
                        endRow,
                        [i],
                        expand: expand ? expand : i < rowIndexes.Max().Value,
                        routes.ToArray());
                }
                else
                {
                    (startRow, endRow) = PopulateSectionDataRoute(worksheet,
                        sectionObject,
                        data,
                        startRow,
                        endRow,
                        path.FirstOrDefault(),
                        parentSection ?? section,
                        routes,
                        i,
                        expand: expand ? expand : i < rowIndexes.Max().Value,
                        path.Where(p => p != path.FirstOrDefault()).ToArray());
                }
            }

            return (startRow, endRow);
        }

        private (int startRow, int endRow) PopulateSectionData(ExcelWorksheet worksheet, Section sectionObject, DocumentParameterViewModel section, int startRow, int endRow, List<int?> rowIndexes = null, bool expand = false, params Route[] routes)
        {
            int sectionRowCount = endRow - startRow;

            int originalStartRow = startRow;
            int originalEndRow = endRow;

            switch (section.ParamDef.Grouping)
            {
                case GroupingType.Table:
                    {
                        rowIndexes ??= section.Children
                            .SelectMany(c => c.Children.Select(cc => cc.RowIndex ?? cc.Index))
                            .Distinct()
                            .OrderBy(c => c)
                            .ToList();

                        for (int i = rowIndexes.Min().Value; i <= rowIndexes.Max().Value; i++)
                        {
                            for (int row = startRow; row < endRow; row++)
                            {
                                for (int col = sectionObject.StartCol; col <= sectionObject.EndCol; col++)
                                {
                                    if (col == sectionObject.StartCol)
                                    {
                                        worksheet.Cells[row, col, row, col].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                                    }
                                    if (col == sectionObject.EndCol)
                                    {
                                        worksheet.Cells[row, col, row, col].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                                    }

                                    var cellValue = worksheet.Cells[row, col].Text;
                                    var match = Regex.Match(cellValue, @"\{(.+?)\}.*", RegexOptions.IgnoreCase | RegexOptions.Singleline);

                                    if (match.Success)
                                    {
                                        string placeholder = match.Groups[1].Value;
                                        string replacement = null;

                                        if (routes.Length > 0)
                                        {
                                            replacement = Helpers.Tools.GetReplacementValue(placeholder, section, index: i, routes: routes);
                                        }
                                        else
                                        {
                                            replacement = Helpers.Tools.GetReplacementValue(placeholder, section, i);
                                        }
                                        if (!string.IsNullOrEmpty(replacement))
                                        {
                                            worksheet.Cells[row, col].Value = Regex.Replace(cellValue, @"\{(.+?)\}", replacement);
                                        }
                                        else
                                        {
                                            worksheet.Cells[row, col].Value = string.Empty;
                                        }
                                    }
                                }
                            }
                            startRow += sectionRowCount;
                            endRow += sectionRowCount;
                            if (i + 1 <= rowIndexes.Max().Value || expand)
                            {
                                worksheet.InsertRow(startRow, sectionRowCount);
                                worksheet.Cells[startRow, 1, endRow, sectionObject.EndCol].Value = sectionObject.Cells;

                                foreach (var mergeAddress in sectionObject.MergedCells)
                                {
                                    ExcelAddress newMergeAddess = null;

                                    var mergeRowOffsetStart = mergeAddress.Start.Row - sectionObject.StartRow;
                                    var mergeRowOffsetEnd = mergeAddress.End.Row - sectionObject.StartRow;

                                    newMergeAddess = new ExcelAddress(startRow + mergeRowOffsetStart,
                                           mergeAddress.Start.Column,
                                           startRow + mergeRowOffsetEnd,
                                           mergeAddress.End.Column);

                                    try
                                    {
                                        worksheet.Cells[newMergeAddess.Address].Merge = true;
                                    }
                                    catch
                                    {
                                    }
                                }
                                for (int j = sectionObject.StartRow; j <= sectionObject.EndRow; j++)
                                {
                                    for (int newCol = sectionObject.StartCol; newCol <= sectionObject.EndCol; newCol++)
                                    {
                                        var cell = worksheet.Cells[startRow + (j - sectionObject.StartRow), newCol, startRow + (j - sectionObject.StartRow), newCol];
                                        var sectionCell = sectionObject.CellStyles.FirstOrDefault(c => c.Row == j && c.Col == newCol);

                                        if (cell is null) continue;

                                        Helpers.Tools.ApplyStyles(cell, sectionCell, j == sectionObject.StartRow, j == sectionObject.EndRow);
                                    }
                                }
                            }
                        }
                    }
                    break;

                case GroupingType.Cluster:
                    for (int row = startRow; row < endRow; row++)
                    {
                        for (int col = sectionObject.StartCol; col <= sectionObject.EndCol; col++)
                        {
                            if (col == sectionObject.StartCol)
                            {
                                worksheet.Cells[row, col, row, col].Style.Border.Left.Style = ExcelBorderStyle.Thick;
                            }
                            if (col == sectionObject.EndCol)
                            {
                                worksheet.Cells[row, col, row, col].Style.Border.Right.Style = ExcelBorderStyle.Thick;
                            }

                            var cellValue = worksheet.Cells[row, col].Text;
                            var match = Regex.Match(cellValue, @"\{(.+?)\}");

                            if (match.Success)
                            {
                                string placeholder = match.Groups[1].Value;
                                string replacement = Helpers.Tools.GetReplacementValue(placeholder, section);

                                if (!string.IsNullOrEmpty(replacement))
                                {
                                    worksheet.Cells[row, col].Value = Regex.Replace(cellValue, @"\{(.+?)\}", replacement);
                                }
                                else
                                {
                                    worksheet.Cells[row, col].Value = string.Empty;
                                }
                            }
                        }
                    }
                    //startRow += sectionRowCount;
                    //endRow += sectionRowCount;
                    break;

                default:
                    break;
            }
            var topRow = worksheet.Cells[originalStartRow, sectionObject.StartCol, originalStartRow, sectionObject.EndCol];
            topRow.Style.Border.Top.Style = ExcelBorderStyle.Medium;

            var bottomRow = worksheet.Cells[endRow, sectionObject.StartCol, endRow, sectionObject.EndCol];
            bottomRow.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            return (startRow, endRow);
        }


        private void SavePackage(ExcelPackage package)
        {
            try
            {
                package.Save();
            }
            catch (Exception ex)
            {
                _logger.LogError($"{DateTime.Now} Excel Generator: Save Package Exception: {ex}");
            }
        }
    }
}
