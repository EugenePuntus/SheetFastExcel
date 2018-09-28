using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Kapral.FastExcel
{
    public class ReceivedDataEventArgs : EventArgs
    {
        public ReceivedDataEventArgs(Range data)
        {
            Data = data;
        }

        public Range Data { get; private set; }
    }

    public interface ISheetFastExcel
    {
        //save data
        void SaveData(object[,] data);
        void SaveData<T>(IEnumerable<T> data);
        void SaveData(object[,] data, int offSetRow);

        string Name { get; }

        int RowsCount { get; }
        int ColumnsCount { get; }

        //get value
        object GetCellValue(int row, int col);
        string GetString(int row, int col);
        DateTime GetDateTime(int row, int col);
        DateTime GetDateTime(int row, int col, IFormatProvider provider);
        double GetDouble(int row, int col);
        double GetDoubleOrDefault(int row, int col, double defaultValue);
        decimal GetDecimal(int row, int col);
        decimal GetDecimalOrDefault(int row, int col, decimal defaultValue);
        int GetInt(int row, int col);
        int GetIntOrDefault(int row, int col, int defaultValue);

        //check
        bool IsEmptyCell(int row, int col);
        bool IsNotEmptyCell(int row, int col);
        bool IsSameStrings(int row, int col, string compareString);
        bool IsContains(int row, int col, string compareString);

        //actions
        event EventExcelRange BeforeSaving;
    }
}
