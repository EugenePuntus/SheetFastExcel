# SheetFastExcel
Class for fast reading and writing Excel

## How to use read Excel
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

## How to use write Excel
```c#

[IgnoreHeader] //if it's used header name not visible on excell sheet
public class ExampleClass
{
    [Column("№")]
    public int Key { get; set; }    
    
    [Column("NameColumn")]
    public string Name { get; set; }

    [Column("HeaderDate","dd.MM.yyyy")] //header name and format data in excel cell
    public DateTime ExamleDate { get; set; }
}

public void SaveToExcel()
{
    using (var excel = new FastExcel())
    {
        var elements = new List<ExampleClass>() { ... };
        var sheet = excel.AddNewSheet("nameSheetExcel");
        sheet.BeforeSaving += UseDefaultStyle; // apply style befare save in excell
        sheet.SaveData(elements);        
        
        var array = new object[,] { ... };
        var secondSheet = excel.AddNewSheet("SecondSheetExcell");
        secondSheet.SaveData(array);

        excel.GenerateAndOpen();
    }
}

private void UseDefaultStyle(Range range)
{
    range.Font.Bold = false;
    range.Font.Italic = false;
    range.Font.Size = 10;
    range.Font.ColorIndex = 1;
    range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
    range.VerticalAlignment = XlVAlign.xlVAlignCenter;

    var borders = range.Borders;
    borders[XlBordersIndex.xlEdgeBottom].Color = borders[XlBordersIndex.xlEdgeLeft].Color = borders[XlBordersIndex.xlEdgeTop].Color = borders[XlBordersIndex.xlInsideHorizontal].Color = borders[XlBordersIndex.xlInsideVertical].Color = borders[XlBordersIndex.xlEdgeRight].Color = 0;
    borders[XlBordersIndex.xlEdgeBottom].LineStyle = borders[XlBordersIndex.xlEdgeLeft].LineStyle = borders[XlBordersIndex.xlEdgeTop].LineStyle = borders[XlBordersIndex.xlInsideHorizontal].LineStyle = borders[XlBordersIndex.xlInsideVertical].LineStyle = borders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
    borders[XlBordersIndex.xlEdgeBottom].Weight = borders[XlBordersIndex.xlEdgeLeft].Weight = borders[XlBordersIndex.xlEdgeTop].Weight = borders[XlBordersIndex.xlInsideHorizontal].Weight = borders[XlBordersIndex.xlInsideVertical].Weight = borders[XlBordersIndex.xlEdgeRight].Weight = 2;
}

```
