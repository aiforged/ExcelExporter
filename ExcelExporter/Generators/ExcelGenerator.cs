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
        private Models.IConfig _config;
        private int _iteration = 0;
        private Stack<Section> _templateStack = new Stack<Section>();

        private const string SECTIONPATTERN = @"\{([a-zA-Z\:\|\&\.]+)\s(start|end)\}";

        public ExcelGenerator(ILogger logger, Models.IConfig config) 
        { 
            this._logger = logger;
            this._config = config;
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
                var sectionData = data.FirstOrDefault(d => d.ParamDef.Name.Equals(_config.MasterParamDefName))?.Children;

                //First pass we focus on iterating through te master data
                foreach (var child in sectionData)
                {
                    currentRow = ProcessSections(worksheet, child, currentRow);

                    if (data.ToList().IndexOf(child) == data.Count - 1) break;
                    
                    _iteration++;
                    if (_iteration < sectionData.Count())
                    {
                        CloneTemplate(worksheet, currentRow);
                    }
                }

                //Last pass we fill in any data that's not included in any sections.
                PopulateNonSectionData(worksheet, data);

                SavePackage(package);
            }
        }

        private int ParseTemplate(ExcelWorksheet worksheet, int startRow)
        {
            string prevSectionName = null;
            int finalEndRow = startRow;
            _templateStack = new Stack<Section>();

            for (int row = startRow; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text;
                    var match = Regex.Match(cellValue, SECTIONPATTERN);

                    if (match.Success)
                    {
                        prevSectionName = HandleSectionMatch(worksheet, row, match, prevSectionName);
                        finalEndRow = row;
                        col = worksheet.Dimension.End.Column; // Force next row
                    }
                }
            }
            return finalEndRow;
        }

        private string HandleSectionMatch(ExcelWorksheet worksheet, int row, Match match, string prevSectionName)
        {
            string sectionName = match.Groups[1].Value;
            string sectionType = match.Groups[2].Value;

            if (sectionType == "start")
            {
                AddSectionStart(worksheet, row + 1, sectionName, prevSectionName);
                return sectionName;
            }
            else if (sectionType == "end" && _templateStack.Count > 0)
            {
                CompleteSectionEnd(worksheet, row - 1, sectionName);
                return null;
            }
            return null;
        }

        private void AddSectionStart(ExcelWorksheet worksheet, int row, string sectionName, string prevSectionName)
        {
            int startColumn = FindStartColumn(worksheet, row);
            int endColumn = FindEndColumn(worksheet, row);

            _templateStack.Push(new Section(sectionName, prevSectionName, row, startColumn, endColumn));
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
                section.Cells = worksheet.Cells[section.StartRow, section.StartCol, row, section.EndCol].Value;
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

                if (section.ParentName is not null)
                {
                    previousSection = section;
                    continue;
                }
                (int currentStartRow, int currentEndRow, rowExpansion) = ProcessSection(worksheet, data, startRow, rowExpansion, section, previousSection);

                //Clear the section header
                worksheet.Cells[section.StartRow - 2 + startRow, section.StartCol, section.StartRow - 2 + startRow, section.EndCol].Value = string.Empty;
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
            worksheet.Cells[startRow, section.StartCol, endRow, section.EndCol].Value = section.Cells;

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
            int sectionRowCount = endRow - startRow;
            int totalRowExpansion = 0;
            //worksheet.DeleteRow(endRow, 1);
            //worksheet.DeleteRow(startRow, 1);

            if (sectionObject.Name.Contains(":") || sectionObject.Name.Contains("&"))
            {
                var paths = sectionObject.Name.Split("&", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

                (startRow, endRow, int rowExpansion) = PopulateSectionDataRoute(worksheet, sectionObject, data, startRow, endRow, paths.FirstOrDefault(), parentSection: null, routes: null, index: null, expand: false, paths.Where(r => r != paths.FirstOrDefault()).ToArray());

                totalRowExpansion += rowExpansion;
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
                    (startRow, endRow, int rowExpansion) = PopulateSectionData(worksheet, sectionObject, section, startRow, endRow);

                    totalRowExpansion += rowExpansion;
                }
            }

            return (startRow, endRow, totalRowExpansion);
        }

        private (int startRow, int endRow, int rowExpansion) PopulateSectionDataRoute(ExcelWorksheet worksheet, Section sectionObject, DocumentParameterViewModel data, int startRow, int endRow, string route, DocumentParameterViewModel parentSection = null, List<Route> routes = null, int? index = null, bool expand = false, params string[] path)
        {
            var section = Helpers.Tools.GetParameterRoute(route.Split("|").FirstOrDefault(), data, index: index, checkColumnChildren: true, routes: routes?.ToArray());

            if (section is null) return (startRow, endRow, 0);
            var rowIndexes = section.Children
                        .SelectMany(c => c.Children.Select(cc => cc.RowIndex ?? cc.Index))
                        .Distinct()
                        .OrderBy(c => c)
                        .ToList();

            if (!rowIndexes.Any()) return (startRow, endRow, 0);
            int rowStep = 0;
            int totalRowExpansion = 0;

            for (int i = 0; i <= rowIndexes.Max().Value; i++)
            {
                totalRowExpansion += rowStep;
                int rowExpansion = 0;
                routes ??= new List<Route>();

                string[] routePipe = route.Split("|", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);

                foreach (string routePiece in routePipe)
                {
                    string[] routeSplit = routePiece.Split(":", StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                    var depth = routeSplit.Count() - 1;
                    var lastDef = routeSplit.LastOrDefault();
                    var foundRoute = routes.FirstOrDefault(r => r.DefName == lastDef);

                    if (foundRoute is null)
                    {
                        routes.Add(new Route()
                        {
                            DefName = lastDef,
                            Depth = depth,
                            Index = index,
                        });
                    }
                    else
                    {
                        foundRoute.Index = index;
                    }
                }

                if (path is null || path.Length == 0)
                {
                    (startRow, endRow, rowExpansion) = PopulateSectionData(worksheet,
                        sectionObject,
                        parentSection ?? data,
                        startRow,
                        endRow,
                        route: route,
                        rowIndexes: [i],
                        expand: expand ? expand : i < rowIndexes.Max().Value,
                        routes.ToArray());
                }
                else
                {
                    (startRow, endRow, rowExpansion) = PopulateSectionDataRoute(worksheet,
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
                totalRowExpansion += rowExpansion;
            }

            return (startRow, endRow, totalRowExpansion);
        }

        private (int startRow, int endRow, int rowExpansion) PopulateSectionData(ExcelWorksheet worksheet, 
            Section sectionObject, 
            DocumentParameterViewModel section, 
            int startRow, 
            int endRow, 
            string route = null, 
            List<int?> rowIndexes = null, 
            bool expand = false, 
            params Route[] routes)
        {
            int originalStartRow = startRow;
            int originalEndRow = endRow;
            int rowExpansion = 0;
            int sectionRowCount = (endRow - startRow) + 1;

            switch (section.ParamDef.Grouping)
            {
                case GroupingType.Table:
                    {
                        rowIndexes ??= section.Children
                            .SelectMany(col => col
                                .Children
                                .Select(cell => cell.RowIndex ?? cell.Index))
                            .Distinct()
                            .OrderBy(i => i)
                            .ToList();

                        if (rowIndexes is null || rowIndexes.Count == 0) return (startRow, endRow, startRow - originalStartRow);
                        for (int i = rowIndexes.Min().Value; i <= rowIndexes.Max().Value; i++)
                        {
                            int sectionRowExpansion = 0;

                            for (int row = startRow; row <= endRow; row++)
                            {
                                for (int col = sectionObject.StartCol; col <= sectionObject.EndCol; col++)
                                {
                                    if (col == sectionObject.StartCol)
                                    {
                                        worksheet.Cells[row, col, row, col].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                                    }
                                    if (col == sectionObject.EndCol)
                                    {
                                        worksheet.Cells[row, col, row, col].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                                    }

                                    string cellValue = worksheet.Cells[row, col].Text;
                                    Match match = Regex.Match(cellValue, @"\{(.+?)\}.*", RegexOptions.IgnoreCase | RegexOptions.Singleline);

                                    if (match.Success)
                                    {
                                        string placeholder = match.Groups[1].Value;
                                        string replacement = null;

                                        //Check if section placeholder and clear values:
                                        Match sectionMatch = Regex.Match(cellValue, SECTIONPATTERN);

                                        if (sectionMatch.Success)
                                        {
                                            if (sectionMatch.Groups[2].Value != "end")
                                            {
                                                Section innerSection = GetSectionObject(sectionMatch);

                                                if (innerSection != sectionObject && innerSection.ParentName == sectionObject.Name)
                                                {
                                                    int innerStartRow = row;
                                                    int innerEndRow = row + (innerSection.EndRow - innerSection.StartRow);

                                                    string startingRoute = innerSection.Name;
                                                    List<string> comboPath = new List<string>();
                                                    List<string> path = new List<string>();

                                                    foreach (string combo in startingRoute.Split("&", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                                                    {
                                                        foreach (string part in combo.Split(":"))
                                                        {
                                                            if (!section.ParamDef.Name.Equals(part))
                                                            {
                                                                path.Add(part);
                                                            }
                                                        }
                                                        if (path.All(string.IsNullOrEmpty)) continue;
                                                        comboPath.Add(string.Join(":", path.Where(p => !string.IsNullOrEmpty(p))));
                                                    }

                                                    string newRoute = string.Join("&", comboPath);

                                                    (int currentStartRow, int currentEndRow, rowExpansion) = PopulateSectionDataRoute(worksheet,
                                                        sectionObject: innerSection,
                                                        data: section,
                                                        startRow: innerStartRow, 
                                                        endRow: innerEndRow, 
                                                        route: newRoute, 
                                                        routes: routes.ToList(), 
                                                        index: i);
                                                    //(int currentStartRow, int currentEndRow, rowExpansion)  = PopulateSectionData(worksheet,
                                                    //    innerSection,
                                                    //    section,
                                                    //    startRow,
                                                    //    endRow,
                                                    //    [i],
                                                    //    expand: expand ? expand : i < rowIndexes.Max().Value,
                                                    //    routes.ToArray());

                                                    endRow += rowExpansion + 1;
                                                    sectionRowExpansion += rowExpansion + 1;
                                                    row = currentEndRow + 1;

                                                    worksheet.Cells[innerStartRow, col].Value = string.Empty;
                                                    worksheet.Cells[row, col].Value = string.Empty;
                                                }
                                            }
                                        }
                                        else if (routes.Length > 0)
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
                            startRow += (sectionRowCount + sectionRowExpansion);
                            endRow += (sectionRowCount);
                            if (i + 1 <= rowIndexes.Max().Value || expand)
                            {
                                worksheet.InsertRow(startRow, sectionRowCount);

                                //var middleElements = Helpers.Tools.GetArraySection((object[,])sectionObject.Cells, 0, sectionRowCount, 0, sectionObject.EndCol);

                                worksheet.Cells[startRow, sectionObject.StartCol, endRow, sectionObject.EndCol].Value = sectionObject.Cells;
                                var rowOffset = startRow - sectionObject.StartRow;

                                foreach (var mergeAddress in sectionObject.MergedCells)
                                {
                                    ExcelAddress newMergeAddess = null;

                                    newMergeAddess = new ExcelAddress(mergeAddress.Start.Row + rowOffset,
                                           mergeAddress.Start.Column,
                                           mergeAddress.End.Row + rowOffset,
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
                                worksheet.Cells[row, col, row, col].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                            }
                            if (col == sectionObject.EndCol)
                            {
                                worksheet.Cells[row, col, row, col].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                            }

                            var cellValue = worksheet.Cells[row, col].Text;
                            var match = Regex.Match(cellValue, @"\{(.+?)\}");

                            if (match.Success)
                            {
                                string placeholder = match.Groups[1].Value;
                                string replacement = null;

                                //Check if section placeholder and clear values:
                                var sectionMatch = Regex.Match(cellValue, SECTIONPATTERN);

                                if (sectionMatch.Success)
                                {
                                    if (sectionMatch.Groups[2].Value != "end")
                                    {
                                        Section innerSection = GetSectionObject(sectionMatch);

                                        if (innerSection != sectionObject && innerSection.ParentName == sectionObject.Name)
                                        {
                                            var innerStartRow = row;
                                            var innerEndRow = row + (innerSection.EndRow - innerSection.StartRow);

                                            (int currentStartRow, int currentEndRow, rowExpansion) = ProcessSection(worksheet, section, startRow - sectionObject.StartRow + 1, rowExpansion, innerSection, sectionObject);

                                            endRow += rowExpansion;
                                            row = currentEndRow - 1;

                                            worksheet.Cells[innerStartRow, col].Value = string.Empty;
                                            worksheet.Cells[row - 1, col].Value = string.Empty;
                                        }
                                    }
                                }
                                else
                                {
                                    replacement = Helpers.Tools.GetReplacementValue(placeholder, section);
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
                    //startRow += sectionRowCount;
                    //endRow += sectionRowCount;
                    break;

                default:
                    break;
            }

            if (sectionObject.ParentName != null) return (startRow, endRow, startRow - originalStartRow);

            var topRow = worksheet.Cells[originalStartRow, sectionObject.StartCol, originalStartRow, sectionObject.EndCol];
            topRow.Style.Border.Top.Style = ExcelBorderStyle.Medium;

            var bottomRow = worksheet.Cells[endRow, sectionObject.StartCol, endRow, sectionObject.EndCol];
            bottomRow.Style.Border.Bottom.Style = ExcelBorderStyle.Medium;

            return (startRow, endRow, (endRow - startRow) - (originalEndRow - originalStartRow));
        }

        private void PopulateNonSectionData(ExcelWorksheet worksheet, ICollection<DocumentParameterViewModel> data)
        {
            if (worksheet is null) return;
            if (data is null) return;

            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
            {
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text;
                    var match = Regex.Match(cellValue, @"\{(.+?)\}.*", RegexOptions.IgnoreCase | RegexOptions.Singleline);

                    if (match.Success)
                    {
                        string placeholder = match.Groups[1].Value;
                        string replacement = null;

                        //Check if section placeholder and clear values:
                        var sectionMatch = Regex.Match(cellValue, SECTIONPATTERN);

                        if (sectionMatch.Success)
                        {
                            //Do nothing
                        }
                        else
                        {
                            foreach (var docParam in data.Where(d => !d.ParamDef.Name.Equals(_config.MasterParamDefName)))
                            {
                                replacement = Helpers.Tools.GetReplacementValue(placeholder, docParam);
                                if (!string.IsNullOrEmpty(replacement)) break;
                            }
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
        }

        private Section GetSectionObject(Match match)
        {
            string sectionName = match.Groups[1].Value;

            return _templateStack.FirstOrDefault(s => s.Name == sectionName);
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
