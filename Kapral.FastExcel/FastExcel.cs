using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Kapral.FastExcel
{
    public class FastExcel : IDisposable
    {
        private readonly Microsoft.Office.Interop.Excel.Application _app;
        private readonly Microsoft.Office.Interop.Excel.Workbook _workBook;

        public IEnumerable<SheetFastExcel> Sheets
        {
            get
            {
                for (int i = 1; i <= _workBook.Worksheets.Count; i++)
                {
                    var sheet = _workBook.Worksheets[i] as Microsoft.Office.Interop.Excel.Worksheet;
                    yield return new SheetFastExcel(sheet);
                }
            }
        }

        public FastExcel(string fileName)
        {
            _app = new Microsoft.Office.Interop.Excel.Application();
            _workBook = _app.Workbooks.Add(fileName);
        }

        public void Dispose()
        {
            try
            {
                if (_app != null)
                {
                    _app.DisplayAlerts = false;
                    _app.Quit();
                }
            }
            catch { }
        }
    }
}
