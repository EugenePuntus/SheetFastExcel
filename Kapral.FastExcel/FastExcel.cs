using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace Kapral.FastExcel
{
    public class FastExcel : IFastExcel, IDisposable
    {
        private readonly Application _app;
        private readonly Workbook _workBook;

        public IEnumerable<ISheetFastExcel> Sheets
        {
            get
            {
                for (int i = 1; i <= _workBook.Worksheets.Count; i++)
                {
                    var sheet = _workBook.Worksheets[i] as Worksheet;
                    yield return new SheetFastExcel(sheet);
                }
            }
        }

        public FastExcel(string filePath)
        {
            _app = new Application();
            _workBook = _app.Workbooks.Add(filePath);
        }

        public FastExcel()
        {
            _app = new Application();
            _workBook = _app.Workbooks.Add();
        }

        public void GenerateAndOpen()
        {
            _app.Visible = true;
        }

        /// <summary>
        /// Adding a new sheet to the document
        /// </summary>
        /// <param name="nameSheet">The size must not exceed 32 characters</param>
        /// <returns></returns>
        public ISheetFastExcel AddNewSheet(string nameSheet)
        {
            //var workSheet = (Excel.Worksheet)objExcel.Worksheets.Add(Type.Missing, objExcel.Worksheets[objExcel.Worksheets.Count], 1, XlSheetType.xlWorksheet);
            var workSheet = (Worksheet) _app.Worksheets.Add(System.Type.Missing, _app.Worksheets[_app.Worksheets.Count], 1, XlSheetType.xlWorksheet);
            workSheet.Name = nameSheet;

            return new SheetFastExcel(workSheet);
        }

        public void Dispose()
        {
            try
            {
                if (_app != null && _app.Visible != true)
                {
                    _app.DisplayAlerts = false;
                    _app.Quit();
                }
            }
            catch { }
        }
    }
}
