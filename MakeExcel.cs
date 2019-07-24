using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml; //add "EPPlus.Core" Nuget

namespace Utility
{
    public static class MakeExcel
    {
        public static ExcelData GetExcel<T>(this IList<T> list, string fileName, bool rightToLeft = true)
        {
            return typeof(T) == typeof(Dictionary<string, object>) ? GetExcelFromDictionary((IList<Dictionary<string, object>>)list, fileName, rightToLeft) : GetExcelFromClass(list, fileName, rightToLeft);
        }

        private static ExcelData GetExcelFromClass<T>(this ICollection<T> list, string fileName, bool rightToLeft = true)
        {
            var entityType = typeof(T);
            var dataTable = new DataTable(entityType.Name);
            var properties = TypeDescriptor.GetProperties(entityType);

            if (list.Count <= 0)
                return DataTableToExcelData(dataTable, fileName, rightToLeft);

            foreach (PropertyDescriptor prop in properties)
                if (prop.PropertyType == typeof(List<string>))
                {
                    var maxColumn = list.Max(a => ((List<string>)prop.GetValue(a) ?? throw new InvalidOperationException()).Count);
                    for (var i = 0; i < maxColumn; i++)
                        dataTable.Columns.Add($"{GetDisplayName<T>(prop)} [{i + 1}]",
                            Nullable.GetUnderlyingType(typeof(string)) ?? typeof(string));
                }
                else
                {
                    dataTable.Columns.Add(GetDisplayName<T>(prop),
                        Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }

            foreach (var item in list)
            {
                var row = dataTable.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    if (prop.PropertyType == typeof(List<string>))
                    {
                        var maxColumn = list.Max(a => ((List<string>)prop.GetValue(a) ?? throw new InvalidOperationException()).Count);
                        var values = (List<string>)prop.GetValue(item);

                        for (var i = 0; i < maxColumn; i++)
                            row[$"{GetDisplayName<T>(prop)} [{i + 1}]"] =
                                values != null && values.Count > i ? values[i] : string.Empty;
                    }
                    else
                    {
                        row[GetDisplayName<T>(prop)] = prop.GetValue(item) ?? DBNull.Value;
                    }

                dataTable.Rows.Add(row);
            }

            return DataTableToExcelData(dataTable, fileName, rightToLeft);
        }

        private static ExcelData GetExcelFromDictionary(this IList<Dictionary<string, object>> dictionaryList, string fileName, bool rightToLeft = true)
        {
            var entityType = dictionaryList.GetType();
            var dataTable = new DataTable(entityType.Name);
            TypeDescriptor.GetProperties(entityType);
            if (!dictionaryList.Any() || dictionaryList.FirstOrDefault() == null)
                return DataTableToExcelData(dataTable, fileName, rightToLeft);

            foreach (var (key, value) in dictionaryList.FirstOrDefault()) dataTable.Columns.Add(key, value.GetType());

            foreach (var item in dictionaryList)
            {
                var row = dataTable.NewRow();
                foreach (var (key, value) in item) row[key] = value ?? DBNull.Value;

                dataTable.Rows.Add(row);
            }
            return DataTableToExcelData(dataTable, fileName, rightToLeft);
        }

        private static ExcelData DataTableToExcelData(DataTable dataTable, string fileName, bool rightToLeft = true)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Excel");
                package.Workbook.Worksheets["Excel"].View.RightToLeft = rightToLeft;

                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);
                //for (int col = 1; col < dataTable.Columns.Count + 1; col++)
                //{
                //    worksheet.Column(col).AutoFit(5, 100);
                //}

                return new ExcelData
                {
                    FileContents = package.GetAsByteArray(),
                    ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    FileDownloadName = $"{fileName}.xlsx"
                };
            }
        }

        private static string GetDisplayName<T>(MemberDescriptor prop)
        {
            return typeof(T)
                .GetProperty(prop.Name)
                .GetCustomAttribute(typeof(DisplayAttribute)) is DisplayAttribute data
                ? data.Name
                : prop.DisplayName ?? prop.Name;
        }

        /// <summary>
        /// GEt class list from excel file
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="file">Excel File</param>
        /// <param name="outPutModel"></param>
        /// <param name="excelFieldPair">key-value pair - key: Excel Field, Value: Class Field Name</param>
        /// <returns></returns>
        public static void GetClass<T>(this IFormFile file, List<T> outPutModel, Dictionary<string, string> excelFieldPair)
        {
            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                stream.Seek(0, SeekOrigin.Begin);

                using (var package = new ExcelPackage(stream))
                {
                    var sheet = package.Workbook.Worksheets[1];

                    var excelFields = new Dictionary<int, string>();

                    for (var j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                    {
                        excelFields.Add(j, sheet.Cells[1, j].Value.ToString());
                    }

                    for (var i = sheet.Dimension.Start.Row + 1; i <= sheet.Dimension.End.Row; i++)
                    {
                        var item = (T)Activator.CreateInstance(typeof(T));
                        for (var j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                        {
                            var data = sheet.Cells[i, j].Value;

                            item.SetClassValue(excelFieldPair.FirstOrDefault(a => a.Key == excelFields[j]).Value, data);
                        }
                        outPutModel.Add(item);
                    }
                }
            }
        }

        private static void SetClassValue<T>(this T item, string itemFieldName, object data)
        {
            if (itemFieldName == null)
                return;

            var entityType = item.GetType();
            var properties = TypeDescriptor.GetProperties(entityType);
            var props = itemFieldName.Contains(".") ? itemFieldName.Split(".") : null;
            if (props != null) itemFieldName = props[0];
            foreach (PropertyDescriptor prop in properties)
                if (prop.Name == itemFieldName)
                {
                    if (props != null)
                    {
                        var instance = Convert.ChangeType(Activator.CreateInstance(prop.PropertyType), prop.PropertyType);
                        instance.SetClassValue(props[1], data);
                        prop.SetValue(item, instance);
                    }
                    else
                        prop.SetValue(item, data is double ? Convert.ChangeType(data, prop.PropertyType) : data);
                    break;
                }
        }
    }

    public class ExcelData
    {
        public byte[] FileContents { get; set; }
        public string ContentType { get; set; }
        public string FileDownloadName { get; set; }
    }
}