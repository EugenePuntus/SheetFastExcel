using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Threading;
using Kapral.FastExcel.FastExcelAttribute;
using Microsoft.Office.Interop.Excel;

namespace Kapral.FastExcel
{
    public class SheetFastExcel : ISheetFastExcel
    {
        private Worksheet workSheet { get; set; }
        private Range usedRange { get; set; }
        private object[,] sheet { get; set; }

        private UsedRangeRow usedRangeRow { get; set; }
        private int offSet;
        private CultureInfo _cultureInfo { get; set; }

        public string Name { get; private set; }
        public int RowsCount { get; private set; }
        public int ColumnsCount { get; private set; }
        public event EventExcelRange BeforeSaving;

        private class UsedRangeRow
        {
            public UsedRangeRow(int start, int end)
            {
                Start = start;
                End = end;
            }

            private int Start { get; set; }
            private int End { get; set; }

            public bool IsContains(int row)
            {
                return Start <= row && End >= row;
            }
        }

        public SheetFastExcel(Worksheet ws)
            : this(ws, new CultureInfo("ru-RU"))
        {
        }

        public SheetFastExcel(Worksheet ws, CultureInfo cultureInfo)
        {
            workSheet = ws;
            Name = workSheet.Name;
            RowsCount = workSheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
            //ColumnsCount = workSheet.UsedRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
            ColumnsCount = workSheet.UsedRange.Columns.Count;
            //var columnsCount2 = ws.UsedRange.Columns[ws.UsedRange.Columns.Count].Column;
            usedRange = workSheet.Cells;
            _cultureInfo = cultureInfo;
        }

        public void SaveData(object[,] data)
        {
            SaveData(data, 0);
        }

        public void SaveData(object[,] data, int offSetRow)
        {
            if(data.Length == 0)
                return;

            var sheetRows = data.GetLength(0);
            var sheetCols = data.GetLength(1);

            var startCell = (Range)workSheet.Cells[1 + offSetRow, 1];
            var endCell = (Range)workSheet.Cells[sheetRows + offSetRow, sheetCols];
            var excelcells = workSheet.Range[startCell, endCell];

            BeforeSaving?.Invoke(excelcells);

            excelcells.Value = data;
        }

        public void SaveData<T>(IEnumerable<T> data)
        {
            SaveData(data, 100);
        }

        public void SaveData<T>(IEnumerable<T> data, int rowsLoadedAtOneTime)
        {
            var properties = new List<PropertyInfo>();
            var headers = new List<string>();

            var value = data.FirstOrDefault();
            var type = value?.GetType() ?? typeof(T);

            var tempPorperty = type
                .GetProperties(
                    BindingFlags.GetProperty |
                    BindingFlags.Instance |
                    BindingFlags.Public |
                    BindingFlags.DeclaredOnly
                );

            var ignoreHeader = type.GetCustomAttributes(typeof(IgnoreHeaderAttribute), true).FirstOrDefault() as IgnoreHeaderAttribute;

            for (var i = 0; i < tempPorperty.Length; i++)
            {
                var attribute = tempPorperty[i].GetCustomAttributes(typeof(ColumnAttribute), true).FirstOrDefault() as ColumnAttribute;
                if (attribute == null) continue;

                headers.Add(attribute.HeaderName);
                properties.Add(tempPorperty[i]);

                var range = workSheet.Columns[properties.Count] as Range;

                if(range == null) continue;

                if (!string.IsNullOrWhiteSpace(attribute.NumberFormat))
                    range.NumberFormat = attribute.NumberFormat;

                if(attribute.ColumnWidth > 0)
                    range.ColumnWidth = attribute.ColumnWidth;

                if (attribute.RowHeight > 0)
                    range.RowHeight = attribute.RowHeight;
            }

            var offSetRow = 0;

            if (ignoreHeader == null && headers.Count > 0)
            {
                var dataHeader = new object[1, headers.Count];
                for (int j = 0; j < headers.Count; j++)
                {
                    dataHeader[0, j] = headers[j];
                }

                SaveData(dataHeader);
                offSetRow++;
            }

            var dataExcel = new object[rowsLoadedAtOneTime, properties.Count];
            int counterRow = 0;

            foreach (var myObject in data)
            {
                int column = 0;

                foreach (var property in properties)
                {
                    var val = property.GetValue(myObject, null).ToString();
                    dataExcel[counterRow, column] = val;
                    column++;
                }

                counterRow++;

                if (counterRow >= rowsLoadedAtOneTime)
                {
                    SaveData(dataExcel, offSetRow);
                    counterRow = 0;
                    offSetRow += rowsLoadedAtOneTime;
                    dataExcel = new object[rowsLoadedAtOneTime, properties.Count];
                }
            }

            var partialDataExcel = new object[counterRow, properties.Count];
            Array.Copy(dataExcel, partialDataExcel, partialDataExcel.Length);
            SaveData(partialDataExcel, offSetRow);
        }

        private void LazyLoadingExcel(int row)
        {
            try
            {
                if (usedRangeRow != null && usedRangeRow.IsContains(row))
                    return;

                if (row > RowsCount)
                    throw new Exception($"Number of columns exceeded allowed ranges. Max {RowsCount}");

                //discharge size
                int sizeRow = 10000;
                //set the offset because array always beginning with 1
                offSet = row - 1;

                int start_row = row;
                int end_row = row + sizeRow > RowsCount ? RowsCount : row + sizeRow;

                int start_col = 1;
                int end_col = ColumnsCount;

                var startCell = (Range)usedRange.Cells[start_row, start_col];
                var endCell = (Range)usedRange.Cells[end_row, end_col];
                var writeRange = usedRange.Range[startCell, endCell];

                sheet = writeRange.Value;
                //var sheet2 = writeRange.Value2;
                usedRangeRow = new UsedRangeRow(start_row, end_row);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        public object GetCellValue(int row, int col)
        {
            try
            {
                if (row > RowsCount || col > ColumnsCount)
                    return null;

                LazyLoadingExcel(row);
                CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
                Thread.CurrentThread.CurrentCulture = _cultureInfo;
                object obj = sheet[row - offSet, col];
                Thread.CurrentThread.CurrentCulture = currentCulture;

                return obj;
            }
            catch (Exception)
            {
                return "";
            }
        }

        public string GetString(int row, int col)
        {
            object cellValue = GetCellValue(row, col);
            string str = string.Empty;
            if (cellValue != null)
            {
                str = cellValue.ToString().Trim();
            }
            return str;
        }

        public bool IsEmptyCell(int row, int col)
        {
            return string.IsNullOrWhiteSpace(GetString(row, col));
        }

        public bool IsNotEmptyCell(int row, int col)
        {
            return !IsEmptyCell(row, col);
        }

        public bool IsSameStrings(int row, int col, string compareString)
        {
            return GetString(row, col).Equals(compareString, StringComparison.InvariantCultureIgnoreCase);
        }

        public bool IsContains(int row, int col, string compareString)
        {
            return GetString(row, col).Contains(compareString);
        }

        public DateTime GetDateTime(int row, int col)
        {
            return GetDateTime(row, col, _cultureInfo);
        }

        public DateTime GetDateTime(int row, int col, IFormatProvider provider)
        {
            object cellValue = GetCellValue(row, col);

            string str = string.Empty;
            if (cellValue != null)
            {
                str = cellValue.ToString();
            }

            DateTime date = DateTime.Parse(str, provider);

            return date;
        }

        public double GetDouble(int row, int col)
        {
            object cellValue = GetCellValue(row, col);
            double result = 0f;
            if (cellValue != null)
            {
                string str = cellValue.ToString();
                double.TryParse(str, out result);
            }

            return Convert.ToDouble(result);
        }

        public double GetDoubleOrDefault(int row, int col, double defaultValue = 0.0)
        {
            double temp;
            var result = double.TryParse(GetString(row, col), out temp);

            return result ? temp : defaultValue;
        }

        public decimal GetDecimal(int row, int col)
        {
            object cellValue = GetCellValue(row, col);
            decimal result = 0m;
            if (cellValue != null)
            {
                string str = cellValue.ToString();
                decimal.TryParse(str, out result);
            }

            return Convert.ToDecimal(result);
        }

        public decimal GetDecimalOrDefault(int row, int col, decimal defaultValue = 0)
        {
            decimal temp;
            var result = decimal.TryParse(GetString(row, col), out temp);

            return result ? temp : defaultValue;
        }

        public int GetInt(int row, int col)
        {
            object cellValue = GetCellValue(row, col);
            int num = 0;
            if (cellValue != null)
            {
                int.TryParse(cellValue.ToString(), out num);
            }

            return num;
        }

        public int GetIntOrDefault(int row, int col, int defaultValue = 0)
        {
            int temp;
            var result = int.TryParse(GetString(row, col), out temp);

            return result ? temp : defaultValue;
        }
    }
}
