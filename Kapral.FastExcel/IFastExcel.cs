using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Kapral.FastExcel
{
    public interface IFastExcel
    {
        IEnumerable<ISheetFastExcel> Sheets { get; }
        ISheetFastExcel AddNewSheet(string nameSheet);
        void GenerateAndOpen();
    }
}
