using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ExcellMapper
{
    public class ExcelMapper<T>
    {
        public List<T> MapExcel(Stream fileStream, int startRow)
        {
            List<T> mappedData = new List<T>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                foreach (var worksheetPart in workbookPart.WorksheetParts.ToList())
                {
                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    foreach (Row row in sheetData.Elements<Row>().Skip(startRow - 1))
                    {
                        T mappedObject = Activator.CreateInstance<T>();
                        int columnIndex = 0;

                        foreach (Cell cell in row.Elements<Cell>())
                        {
                            string columnName = GetColumnNameFromCellReference(cell.CellReference);
                            PropertyInfo property = null;

                            if (columnIndex >= 0)
                            {
                                // If column index is provided, ignore column name and use index-based mapping
                                property = typeof(T).GetProperties()
                                    .FirstOrDefault(p => p.GetCustomAttribute<ExcelColumnAttribute>()?.ColumnIndex-1 == columnIndex);
                            }
                            if(property is null && columnIndex >= 0 )
                            {
                                // If column index is not provided, use column name for mapping
                                property = typeof(T).GetProperties().FirstOrDefault(p =>
                                {
                                    var excelColumnAttribute = p.GetCustomAttribute<ExcelColumnAttribute>();
                                    return excelColumnAttribute != null && excelColumnAttribute.ColumnName == columnName;
                                });
                            }
                            if (property != null)
                            {
                                if (!IsPropertyIgnored(property))
                                {
                                    object cellValue = GetCellValue(cell, workbookPart);
                                    object convertedValue = ConvertValue(cellValue, property.PropertyType);
                                    property.SetValue(mappedObject, convertedValue);
                                }
                            }
                            columnIndex++;
                        }

                        mappedData.Add(mappedObject);
                    }
                }

                
            }

            return mappedData;
        }
        private object ConvertValue(object value, Type targetType)
        {
            if (value == null || targetType.IsInstanceOfType(value))
            {
                return value;
            }

            if (targetType == typeof(int))
            {
                int intValue;
                if (int.TryParse(value.ToString(), out intValue))
                {
                    return intValue;
                }
            }

            // Add additional type conversions as needed...

            return value; // Fallback to returning the original value
        }

        private string GetColumnNameFromCellReference(string cellReference)
        {
            // Extracts the column name from the cell reference (e.g., "A1" -> "A")
            return Regex.Replace(cellReference, "[^A-Z]", string.Empty);
        }

        private object GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            // Retrieve the value from the cell based on its data type
            string cellValue = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int sharedStringIndex = int.Parse(cellValue);
                SharedStringItem sharedStringItem = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                return sharedStringItem.Text?.Text;
            }

            return cellValue;
        }

        private bool IsPropertyIgnored(PropertyInfo property)
        {
            // Check if the property is marked with the IgnoreExcelColumnAttribute and its Ignore property is set to true
            return property.GetCustomAttribute<IgnoreExcelColumnAttribute>()?.Ignore == true;
        }
    }
}
