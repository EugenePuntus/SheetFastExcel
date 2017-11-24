using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Kapral.FastExcel
{
    public class SheetFastExcel
    {
        private Range usedRange { get; set; }
        private object[,] sheet { get; set; }

        private UsedRangeRow usedRangeRow { get; set; }
        private int offSet;
        private CultureInfo _cultureInfo { get; set; }

        public string Name { get; private set; }
        public int RowsCount { get; private set; }
        public int ColumnsCount { get; private set; }

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

        public SheetFastExcel(Worksheet ws) : this(ws, new CultureInfo("ru-RU"))
        {
        }

        public SheetFastExcel(Worksheet ws, CultureInfo cultureInfo)
        {
            Name = ws.Name;
            RowsCount = ws.UsedRange.Rows.Count;
            ColumnsCount = ws.UsedRange.Columns.Count;
            usedRange = ws.UsedRange;
            _cultureInfo = cultureInfo;
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
                //sheet = writeRange.Value2;
                usedRangeRow = new UsedRangeRow(start_row, end_row);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private object GetCellValue(int row, int col)
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
