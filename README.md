# SheetFastExcel
Class for fast reading Excel

## How to use
```c#

// select the input file path and run FastExcel
using (var excel = new FastExcel(filePath))
{
    // pass through all sheets
    foreach (var sheet in excel.Sheets)
    {    
        // get value
        var item1 = sheet.GetString(5, 1);
         
        var column = 100;
        for (int row = 7; row < sheet.RowsCount; row++)
        {
          // possible checking
          if(sheet.IsSameStrings(row, 1, "end"))
            break;
          
          // get value
          var item2 = sheet.GetIntOrDefault(row, column, 0);
          var item3 = sheet.GetDecimal(row, 10);
        }
    }
}

```
