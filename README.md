- Spreadsheet ML library independent from Open XML SDK
- Used to create, open, and modify Excel workbooks
- Uses and packages raw XML according to ECMA-376 (Office Open XML specification)
- Intuitive API supports multiple worksheets, styles, number formats, data validation
- Tested with 166 unit tests and compatible with Excel 2007 - 2016+

# Create
```csharp
var wb = new Workbook(); // Create workbook.

var ws = wb.Worksheets.Add("Sheet1"); // Add worksheet.

var cell = ws.Cell("A1"); // Get cell.

cell.Value = 1000; // Set value.

cell.Style.Fill.Color = Color.Yellow; // Color cell.

cell.Style.NumberFormat.Code = Code.Thou0; // Format cell.

cell.DataValidation.List("1000, 2000"); // Add suggestions.

wb.Save("Book1.xlsx"); // Save workbook.
```

# Open
```csharp
var wb = new Workbook("Book1.xlsx"); // Open workbook.

var ws = wb.Worksheets.First(); // Get worksheet.

if (!ws.IsEmpty) // If not empty,
{
    for (int r = ws.FirstRow; r <= ws.LastRow; r++) // Traverse rows.
    {
        for (int c = ws.FirstCol; c <= ws.LastCol; c++) // Traverse columns.
        {
            var value = ws.Cell(r, c).Value; // Get value.
        }
    }
}
```
