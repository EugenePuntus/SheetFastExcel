using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace Kapral.FastExcel.FastExcelAttribute
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public ColumnAttribute()
        {
        }
        
        public ColumnAttribute(string headerName) : this()
        {
            HeaderName = headerName;
        }

        public ColumnAttribute(string headerName, string numberFormat) : this(headerName)
        {
            NumberFormat = numberFormat;
        }

        public string HeaderName { get; set; }

        /// <summary>
        /// General, 0 and other
        /// </summary>
        public string NumberFormat { get; set; }

        public double ColumnWidth { get; set; }

        public double RowHeight { get; set; }
    }
}
