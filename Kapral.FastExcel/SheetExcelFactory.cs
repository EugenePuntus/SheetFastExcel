using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Kapral.FastExcel
{
    public abstract class SheetExcelFactory
    {
        public virtual ISheetFastExcel Get(Worksheet workSheet)
        {
            return new SheetFastExcel(workSheet);
        }
    }
}
