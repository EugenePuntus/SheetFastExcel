using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Kapral.FastExcel
{
    public interface IFastExcel : IDisposable
    {
        IEnumerable<ISheetFastExcel> Sheets { get; }
        ISheetFastExcel AddNewSheet(string nameSheet);
        void GenerateAndOpen();
        void Save();
        void SaveAs(string filename);
        void SaveCopyAs(string filename);
    }
}
