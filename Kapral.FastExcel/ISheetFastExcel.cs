using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Kapral.FastExcel
{
    public interface ISheetFastExcel
    {
        void SaveData(object[,] data, int dataOffset);
        object GetCellValue(int row, int col);
        string GetString(int row, int col);
        bool IsEmptyCell(int row, int col);
        bool IsNotEmptyCell(int row, int col);
        bool IsSameStrings(int row, int col, string compareString);
        bool IsContains(int row, int col, string compareString);
        DateTime GetDateTime(int row, int col);
        DateTime GetDateTime(int row, int col, IFormatProvider provider);
        double GetDouble(int row, int col);
        double GetDoubleOrDefault(int row, int col, double defaultValue);
        decimal GetDecimal(int row, int col);
        decimal GetDecimalOrDefault(int row, int col, decimal defaultValue);
        int GetInt(int row, int col);
        int GetIntOrDefault(int row, int col, int defaultValue);
    }
}
