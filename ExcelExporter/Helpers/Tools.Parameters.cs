using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

using AIForged.API;

using ExcelExporter.Models;

using LarcAI;

namespace ExcelExporter.Helpers
{
    public static partial class Tools
    {

        public static string GetReplacementValue(string placeholder, DocumentParameterViewModel data, int? index = null, int depth = -1, bool canNextBeSelf = true, params Route[] routes)
        {
            if (string.IsNullOrEmpty(placeholder) || data is null) return string.Empty;
            var combination = placeholder.Split("|", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            StringBuilder stringBuilder = new StringBuilder();

            foreach (var combo in combination)
            {
                if (string.IsNullOrEmpty(combo)) continue;
                var path = $"{combo}";
                var key = path;

                if (combo.IndexOf(":") > -1)
                {
                    key = path.Substring(0, combo.IndexOf(":"));
                }

                if (!key.Equals(path))
                {
                    var paths = path.Split(":").ToList();

                    paths.RemoveAt(0);
                    path = string.Join(":", paths);
                    if (paths.FirstOrDefault() == key)
                    {
                        if (!string.IsNullOrEmpty($"{stringBuilder}"))
                        {
                            stringBuilder.Append("|");
                        }
                        stringBuilder.Append(GetReplacementValue(path, GetParameter(key, data, depth: depth + 1, canBeSelf: canNextBeSelf, routes: routes), index, depth + 1, canNextBeSelf: false, routes: routes));
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty($"{stringBuilder}"))
                        {
                            stringBuilder.Append("|");
                        }
                        stringBuilder.Append(GetReplacementValue(path, GetParameter(key, data, depth: depth + 1, canBeSelf: canNextBeSelf, routes: routes), index, depth + 1, routes: routes));
                    }
                }
                else
                {
                    var docParam = GetParameter(key, data, depth: depth + 1, routes: routes);

                    if (docParam is null) continue;
                    if (docParam.ParamDef.Grouping == GroupingType.Column)
                    {
                        foreach (var param in docParam.Children)
                        {
                            docParam = GetParameter(key, param, index, depth: depth + 1, canBeSelf: true, routes: routes);

                            if (docParam is not null) break;
                        }
                    }

                    if (docParam is null) continue;
                    if (!string.IsNullOrEmpty($"{stringBuilder}"))
                    {
                        stringBuilder.Append("|");
                    }
                    stringBuilder.Append(docParam.Value);
                }
            }

            return stringBuilder.ToString();
        }

        public static DocumentParameterViewModel GetParameterRoute(string route, DocumentParameterViewModel data, int? index = null, bool checkColumnChildren = false, bool canNextBeSelf = true, params Route[] routes)
        {
            if (string.IsNullOrEmpty(route) || data is null) return null;
            DocumentParameterViewModel docParam = null;

            if (string.IsNullOrEmpty(route)) return null;
            var path = $"{route}";
            var key = path;

            if (route.IndexOf(":") > -1)
            {
                key = path.Substring(0, route.IndexOf(":"));
            }

            if (!key.Equals(path))
            {
                path = path.Replace($"{key}:", "");
                if (path.Split(":").FirstOrDefault() == key)
                {
                    docParam = GetParameterRoute(path, GetParameter(key, data, routes: routes), index, checkColumnChildren: checkColumnChildren, canNextBeSelf: false, routes: routes);
                }
                else
                {
                    docParam = GetParameterRoute(path, GetParameter(key, data, routes: routes), index, checkColumnChildren: checkColumnChildren, routes: routes);
                }
            }
            else
            {
                docParam = GetParameter(key, data, index, canBeSelf: canNextBeSelf, routes: routes);

                if (docParam is null) return null;
                if (docParam.ParamDef.Grouping == GroupingType.Column && checkColumnChildren)
                {
                    foreach (var param in docParam.Children)
                    {
                        switch (param.ParamDef.Grouping)
                        {
                            default:
                                docParam = GetParameter(key, param, index: index, canBeSelf: true, routes: routes);
                                break;
                        }

                        if (docParam is not null) break;
                    }
                }
            }

            return docParam;
        }

        public static DocumentParameterViewModel GetParameter(string defName, DocumentParameterViewModel data, int? index = null, int depth = -1, bool canBeSelf = true, params Route[] routes)
        {
            if (string.IsNullOrEmpty(defName)) return null;
            if (data is null) return null;

            if (routes is not null && routes.Length > 0)
            {
                var route = routes.FirstOrDefault(r => r.Depth == depth);

                if (route is not null)
                {
                    if (canBeSelf && (data.ParamDef?.Name?.Equals(defName) ?? false) && (route.Index == null || (data.RowIndex ?? data.Index) == route.Index)) return data;
                    //else if (canBeSelf && (data.ParamDef?.Name?.Equals(defName) ?? false) && (index == null || (data.RowIndex ?? data.Index) == index)) return data;
                }
                else if (canBeSelf && (data.ParamDef?.Name?.Equals(defName) ?? false) && (index == null || (data.RowIndex ?? data.Index) == index)) return data;
            }
            else
            {
                if (canBeSelf && (data.ParamDef?.Name?.Equals(defName) ?? false) && (index == null || (data.RowIndex ?? data.Index) == index)) return data;
            }

            if (data?.Children is null || data.Children.Count() == 0) return null;
            DocumentParameterViewModel docParam = null;

            if (routes is not null && routes.Length > 0)
            {
                var route = routes.FirstOrDefault(r => r.Depth == depth);

                if (route is not null)
                {
                    docParam = data.Children.FirstOrDefault(d => (d.ParamDef?.Name?.Equals(defName) ?? false) && (route.Index == null || (d.RowIndex ?? d.Index) == route.Index));
                }
            }
            else if (canBeSelf)
            {
                docParam = data.Children.FirstOrDefault(d => (d.ParamDef?.Name?.Equals(defName) ?? false) && (index == null || (d.RowIndex ?? d.Index) == index));
            }

            if (docParam is not null)
            {
                if (depth == -1)
                {
                    var tempDocParam = GetParameter(defName, docParam, index, depth, canBeSelf: false, routes);

                    if (tempDocParam is not null) docParam = tempDocParam;
                }
                return docParam;
            }
            foreach (var item in data.Children)
            {
                docParam = GetParameter(defName, item, index, depth, routes: routes);
                if (docParam != null) return docParam;
            }
            return docParam;
        }

        public static List<DocumentParameterViewModel> GetParameters(string defName, DocumentParameterViewModel data, int? index = null)
        {
            List<DocumentParameterViewModel> parameters = new List<DocumentParameterViewModel>();
            if (string.IsNullOrEmpty(defName)) return parameters;
            if (data is null) return parameters;

            if (data.ParamDef?.Name?.Equals(defName) ?? false && (index == null || (data.RowIndex ?? data.Index) == index))
            {
                parameters.Add(data);
            }

            if (data?.Children is null || data.Children.Count() == 0) return parameters;
            foreach (var item in data.Children)
            {
                List<DocumentParameterViewModel> docParams = GetParameters(defName, item);
                parameters.AddRange(docParams);
            }
            return parameters;
        }

        public static DocumentParameterViewModel GetParameter(ICollection<DocumentParameterViewModel> hierarchy, string name)
        {
            if (hierarchy == null || hierarchy.Count == 0 || string.IsNullOrWhiteSpace(name)) return null;
            DocumentParameterViewModel docParam = hierarchy.FirstOrDefault(d => d.ParamDef.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (docParam != null) return docParam;
            foreach (var item in hierarchy)
            {
                docParam = GetParameter(item.Children, name);
                if (docParam != null) return docParam;
            }
            return docParam;
        }

        public static DocumentParameterViewModel GetParameter(ICollection<DocumentParameterViewModel> hierarchy, string name, int index)
        {
            if (hierarchy == null || hierarchy.Count == 0 || string.IsNullOrWhiteSpace(name)) return null;
            DocumentParameterViewModel docParam = hierarchy.FirstOrDefault(d => d.ParamDef.Name.Equals(name, StringComparison.OrdinalIgnoreCase) && (d.Index ?? d.RowIndex) == index);

            if (docParam != null) return docParam;
            foreach (var item in hierarchy)
            {
                docParam = GetParameter(item.Children, name, index);
                if (docParam != null) return docParam;
            }
            return docParam;
        }

        public static DocumentParameterViewModel GetColParameter(ICollection<DocumentParameterViewModel> hierarchy, string name, int index)
        {
            if (hierarchy == null || hierarchy.Count == 0 || string.IsNullOrWhiteSpace(name)) return null;
            DocumentParameterViewModel docParam = hierarchy.FirstOrDefault(d => d.ParamDef.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (docParam?.Children?.Count > 0)
            {
                return GetParameter(docParam.Children, name, index);
            }
            foreach (var item in hierarchy)
            {
                docParam = GetParameter(item.Children, name, index);
                if (docParam != null) return docParam;
            }
            return docParam;
        }

        public static List<DocumentParameterViewModel> GetParameters(ICollection<DocumentParameterViewModel> hierarchy, string name, List<DocumentParameterViewModel> foundList = null)
        {
            if (hierarchy == null || hierarchy.Count == 0 || string.IsNullOrWhiteSpace(name)) return null;
            foundList ??= new List<DocumentParameterViewModel>();
            List<DocumentParameterViewModel> docParams = hierarchy.Where(d => d.ParamDef.Name.Equals(name, StringComparison.OrdinalIgnoreCase))?.ToList();

            if (docParams != null && docParams.Count > 0) foundList.AddRange(docParams);
            foreach (var item in hierarchy)
            {
                GetParameters(item.Children, name, foundList);
            }
            return foundList;
        }

        public static List<DocumentParameterViewModel> GetRowsByFieldValue(ICollection<DocumentParameterViewModel> tableCols, string fieldName, string fieldValue)
        {
            List<DocumentParameterViewModel> foundList = new List<DocumentParameterViewModel>();

            if (tableCols == null || tableCols.Count == 0) return foundList;
            List<int> rows = tableCols
                .FirstOrDefault(c => c.ParamDef.Name.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                .Children
                .Where(c => c.Value.Equals(fieldValue, StringComparison.OrdinalIgnoreCase))
                .Select(c => c.RowIndex ?? c.Index ?? 0)
                .ToList();

            foreach (DocumentParameterViewModel col in tableCols)
            {
                DocumentParameterViewModel parentCol = new DocumentParameterViewModel()
                {
                    ParamDef = col.ParamDef,
                    Children = new ObservableCollection<DocumentParameterViewModel>()
                };

                foreach (var row in rows)
                {
                    DocumentParameterViewModel cell = col.Children.FirstOrDefault(c => (c.RowIndex ?? c.Index) == row);

                    if (cell == null) continue;
                    parentCol.Children.Add(cell);
                }
                foundList.Add(parentCol);
            }

            return foundList;
        }

        public static List<DocumentParameterViewModel> SplitRowsIntoParams(this DocumentParameterViewModel table)
        {
            List<DocumentParameterViewModel> returnData = new List<DocumentParameterViewModel>();

            if (table == null) return returnData;
            List<int> indexes = table.Children.SelectMany(c => c.Children.Select(cc => cc.RowIndex ?? cc.Index ?? -1)).Distinct().ToList();

            foreach (var row in indexes)
            {
                DocumentParameterViewModel rowData = new DocumentParameterViewModel()
                {
                    ParamDef = table.ParamDef,
                    Children = new ObservableCollection<DocumentParameterViewModel>()
                };

                rowData.Children = table.Children
                    .Select(c => LarcAI.Utilities.CloneAs<DocumentParameterViewModel>(c))
                    .ToObservableCollection();

                foreach (DocumentParameterViewModel col in rowData.Children)
                {
                    col.Children = table.Children
                        .SelectMany(c => c.Children.Where(cc => (cc.RowIndex ?? cc.Index) == row && cc.ParamDef.ParentId == col.ParamDef.Id))
                        .ToObservableCollection();
                }

                returnData.Add(rowData);
            }

            return returnData;
        }
    }
}
