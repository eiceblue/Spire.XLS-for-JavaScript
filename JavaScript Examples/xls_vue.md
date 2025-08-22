# Excel Workbook Creation with Five Sheets
## Create an Excel workbook with five sheets and populate them with data
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Add 5 empty sheets to the workbook
workbook.CreateEmptySheets(5);

// Loop through each sheet to populate it with data
for (let i = 0; i < 5; i++) {
  let sheet = workbook.Worksheets.get(i);
  sheet.Name = `Sheet${i}`;
  for (let row = 1; row <= 150; row++) {
    for (let col = 1; col <= 50; col++) {
      sheet.Range.get({
        row: row,
        column: col,
      }).Text = `row${row} col${col}`;
    }
  }
}
```

---

# Spire.XLS JavaScript Excel Creation
## Create an Excel workbook with multiple sheets and populate with data
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Add 5 empty sheets to the workbook
workbook.CreateEmptySheets(5);

// Loop through each sheet to populate it with data
for (let i = 0; i < 5; i++) {
    let sheet = workbook.Worksheets.get(i);
    sheet.Name = `Sheet${i}`;
    for (let row = 1; row <= 150; row++) {
        for (let col = 1; col <= 50; col++) {
            sheet.Range.get({row:row, column:col}).Text = `row${row} col${col}`;
        }
    }
}

// Save the workbook to the specified path
workbook.SaveToFile({fileName: 'CreateAnExcelWithFiveSheet.xlsx', fileFormat: wasmModule.ExcelVersion.Version2010});

// Clean up resources
workbook.Dispose();
```

---

# Creating Multiple Excel Files
## Create fifty Excel workbooks, each containing five worksheets filled with data
```javascript
// Loop to create 50 Excel workbooks, each containing 5 sheets
for (let n = 0; n < 50; n++) {
  // Create a new workbook
  let workbook = wasmModule.Workbook.Create();
  // Add 5 empty sheets to the workbook
  workbook.CreateEmptySheets(5);
  // Fill the worksheets with data
  for (let i = 0; i < 5; i++) {
      let sheet = workbook.Worksheets.get(i);
      sheet.Name = `Sheet${i}`;
      for (let row = 1; row <= 151; row++) {
          for (let col = 1; col <= 51; col++) {
              sheet.Range.get({row:row, column:col}).Text = `row${row} col${col}`;
          }
      }
  }

  // Define the output file name 
  outputFileName = `CreateFiftyExcelFiles_${n + 1}.xlsx`;

  // Save the workbook to the specified path
  workbook.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.ExcelVersion.Version2010});
}
```

---

# Excel Hello World Creation
## Create a simple Excel file with Hello World text
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Clear default worksheets
workbook.Worksheets.Clear();

// Add a new worksheet named "MySheet"
const sheet = workbook.Worksheets.Add("MySheet");

// Set text for the "A1" range
sheet.Range.get("A1").Text = "Hello World";

// Set the column width to auto fit
sheet.Range.get("A1").AutoFitColumns();

// Save the workbook to the specified path
workbook.SaveToFile({fileName: 'HelloWorld.xlsx', version: wasmModule.ExcelVersion.Version2010});
```

---

# Open Existing Excel File
## Load an existing Excel file and perform operations on it
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Load an existing Excel from the virtual file system
workbook.LoadFromFile("templateAz2.xlsx");
    
// Add a new worksheet named "MySheet"
let sheet = workbook.Worksheets.Add("MySheet");

// Set text for the "A1" range
sheet.Range.get("A1").Text = "Hello World";

// Clean up resources
workbook.Dispose();
```

---

# Excel Label Control Addition
## Add a label control to an Excel worksheet
```javascript
// Add a label control
let label = sheet.LabelShapes.AddLabel(10, 2, 30, 200);
label.Text = "This is a Label Control";
```

---

# spire.xls javascript listbox
## add listbox control to excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Add an empty sheet to the workbook
workbook.CreateEmptySheets(1);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Set text for cells
sheet.Range.get("A7").Text = "Beijing";
sheet.Range.get("A8").Text = "New York";
sheet.Range.get("A9").Text = "ChengDu";
sheet.Range.get("A10").Text = "Paris";
sheet.Range.get("A11").Text = "Boston";
sheet.Range.get("A12").Text = "London";

sheet.Range.get("C13").Text = "City :";
sheet.Range.get("C13").Style.Font.IsBold = true;

// Add listbox control
let listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80);
listBox.SelectionType = wasmModule.SelectionType.Single;
listBox.SelectedIndex = 2;
listBox.Display3DShading = true;
listBox.ListFillRange = sheet.Range.get("A7:A12");
```

---

# spire.xls javascript scrollbar control
## add scrollbar control to Excel worksheet
```javascript
// Set a value for range B10
sheet.Range.get("B10").NumberValue = 1;
sheet.Range.get("B10").Style.Font.IsBold = true;

// Add scroll bar control
let scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20);
scrollBar.LinkedCell = sheet.Range.get("B10");
scrollBar.Min = 1;
scrollBar.Max = 150;
scrollBar.IncrementalChange = 1;
scrollBar.Display3DShading = true;
```

---

# Excel Table with Filter
## Add a table with filter functionality to an Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Load an existing Excel from the virtual file system
workbook.LoadFromFile(excelFileName);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Create a List Object named Table
let range = sheet.Range.get({
  row: 1,
  column: 1,
  lastRow: sheet.LastRow,
  lastColumn: sheet.LastColumn,
});
sheet.ListObjects.Create("Table", range);

// Set the BuiltInTableStyle for List object
sheet.ListObjects.get(0).BuiltInTableStyle =
  wasmModule.TableBuiltInStyles.TableStyleLight9;

// Save the workbook to the specified path
workbook.SaveToFile({
  fileName: outputFileName,
  version: wasmModule.ExcelVersion.Version2013,
});
```

---

# spire.xls javascript table
## add total row to table
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Create a table with the data from the specific cell range.
let table = sheet.ListObjects.Create("Table", sheet.Range.get("A1:D4"));

//Display total row.
table.DisplayTotalRow = true;
//Add a total row.
let cols = table.Columns;
cols.get(0).TotalsRowLabel = "Total";
cols.get(1).TotalsCalculation = wasmModule.ExcelTotalsCalculation.Sum;
cols.get(2).TotalsCalculation = wasmModule.ExcelTotalsCalculation.Sum;
cols.get(3).TotalsCalculation = wasmModule.ExcelTotalsCalculation.Sum;
```

---

# Excel Subscript and Superscript Formatting
## Apply subscript and superscript formatting to text in Excel cells
```javascript
// Set text for labels
sheet.Range.get("B2").Text = "This is an example of Subscript:";
sheet.Range.get("D2").Text = "This is an example of Superscript:";

// Set the rtf value of "B3" to "R100-0.06"
let range = sheet.Range.get("B3");
range.RichText.Text = "R100-0.06";

// Create a font. Set the IsSubscript property of the font to "true"
let font = workbook.CreateFont();
font.IsSubscript = true;
font.Color = wasmModule.Color.get_Green();

// Set font for specified range of the text in "B3"
range.RichText.SetFont(4, 8, font);

// Set the rtf value of "D3" to "a2 + b2 = c2"
range = sheet.Range.get("D3");
range.RichText.Text = "a2 + b2 = c2";

// Create a font. Set the IsSuperscript property of the font to "true"
font = workbook.CreateFont();
font.IsSuperscript = true;

// Set font for specified range of the text in "D3"
range.RichText.SetFont(1, 1, font);
range.RichText.SetFont(6, 6, font);
range.RichText.SetFont(11, 11, font);
```

---

# Clone Excel Font Style
## Demonstrates how to clone font styles in Excel cells
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Add the text to the Excel sheet cell range A1
sheet.Range.get("A1").Text = "Text1";

// Set A1 cell range's CellStyle
let style = workbook.Styles.Add("style");
style.Font.FontName = "Calibri";
style.Font.Color = wasmModule.Color.get_Red();
style.Font.Size = 12;
style.Font.IsBold = true;
style.Font.IsItalic = true;
sheet.Range.get("A1").CellStyleName = style.Name;

// Clone the same style for B2 cell range
let csOrieign = style.clone();
sheet.Range.get("B2").Text = "Text2";
sheet.Range.get("B2").CellStyleName = csOrieign.Name;

// Clone the same style for C3 cell range and then reset the font color for the text
let csGreen = style.clone();
csGreen.Font.Color = wasmModule.Color.get_Green();
sheet.Range.get("C3").Text = "Text3";
sheet.Range.get("C3").CellStyleName = csGreen.Name;
```

---

# Excel Cell Range Copying
## Copy a range of cells to another location in a worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Specify a destination range
let cells = sheet.Range.get("G1:H19");

// Copy the selected range to destination range
sheet.Range.get("B1:C19").Copy({ destRange: cells });
```

---

# Spire.XLS JavaScript Copy Data with Style
## Demonstrates how to copy data with style from one range to another in an Excel worksheet
```javascript
// Get a source range (A1:D3)
let srcRange = worksheet.Range.get("A1:D3");

// Create a style object
let style = workbook.Styles.Add("style");

// Specify the font attribute
style.Font.FontName = "Calibri";

// Specify the shading color
style.Font.Color = wasmModule.Color.get_Red();

// Specify the border attributes
style.Borders.get(wasmModule.BordersLineType.EdgeTop).LineStyle =
  wasmModule.LineStyleType.Thin;
style.Borders.get(wasmModule.BordersLineType.EdgeTop).Color =
  wasmModule.Color.get_Blue();
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle =
  wasmModule.LineStyleType.Thin;
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).Color =
  wasmModule.Color.get_Blue();
style.Borders.get(wasmModule.BordersLineType.EdgeLeft).LineStyle =
  wasmModule.LineStyleType.Thin;
style.Borders.get(wasmModule.BordersLineType.EdgeLeft).Color =
  wasmModule.Color.get_Blue();
style.Borders.get(wasmModule.BordersLineType.EdgeRight).LineStyle =
  wasmModule.LineStyleType.Thin;
style.Borders.get(wasmModule.BordersLineType.EdgeRight).Color =
  wasmModule.Color.get_Blue();
srcRange.CellStyleName = style.Name;

// Set the destination range
let destRange = worksheet.Range.get("A12:D14");

// Copy the range data with style
srcRange.Copy(destRange, true, true);
```

---

# Spire.XLS JavaScript Copy Formula Values
## Copy only formula values when copying cells in Excel
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Load an existing Excel from the virtual file system
workbook.LoadFromFile(excelFileName);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Set the copy option
const copyOptions = wasmModule.CopyRangeOptions.OnlyCopyFormulaValue;

// Copy range
sheet.Copy(
  sheet.Range.get("A2:C2"),
  sheet.Range.get("A5:C5"),
  copyOptions
);
```

---

# Excel Nested Group Creation
## Create nested groups in Excel using JavaScript
```javascript
// Set the style
let style = workbook.Styles.Add("style");
style.Font.Color = wasmModule.Color.get_CadetBlue();
style.Font.IsBold = true;

// Set the summary rows to appear above detail rows
sheet.PageSetup.IsSummaryRowBelow = false;

// Insert sample data to cells
sheet.Range.get("A1").Value = "Project plan for project X";
sheet.Range.get("A1").CellStyleName = style.Name;

sheet.Range.get("A3").Value = "Set up";
sheet.Range.get("A3").CellStyleName = style.Name;
sheet.Range.get("A4").Value = "Task 1";
sheet.Range.get("A5").Value = "Task 2";
sheet.Range.get("A4:A5").BorderAround({
  borderLine: wasmModule.LineStyleType.Thin,
});
sheet.Range.get("A4:A5").BorderInside({
  borderLine: wasmModule.LineStyleType.Thin,
});

sheet.Range.get("A7").Value = "Launch";
sheet.Range.get("A7").CellStyleName = style.Name;
sheet.Range.get("A8").Value = "Task 1";
sheet.Range.get("A9").Value = "Task 2";
sheet.Range.get("A8:A9").BorderAround({
  borderLine: wasmModule.LineStyleType.Thin,
});
sheet.Range.get("A8:A9").BorderInside({
  borderLine: wasmModule.LineStyleType.Thin,
});

// Group the rows that you want to group
sheet.GroupByRows(2, 9, false);
sheet.GroupByRows(4, 5, false);
sheet.GroupByRows(8, 9, false);
```

---

# spire.xls javascript table
## create table in Excel worksheet
```javascript
// Add a new List Object to the worksheet
sheet.ListObjects.Create(
  "table",
  sheet.Range.get({ row: 1, column: 1, lastRow: 19, lastColumn: 5 })
);
// Add Default Style to the table
sheet.ListObjects.get(0).BuiltInTableStyle =
  wasmModule.TableBuiltInStyles.TableStyleLight9;
```

---

# Excel Data Sorting
## Sort values in a specific cell range of a worksheet
```javascript
// Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

// Add sorting columns (column 2 and 3 in ascending order)
workbook.DataSorter.SortColumns.Add(2, wasmModule.OrderBy.Ascending);
workbook.DataSorter.SortColumns.Add(3, wasmModule.OrderBy.Ascending);

// Apply sorting to the specified range
workbook.DataSorter.Sort(worksheet.Range.get("A1:E19"));
```

---

# Excel Data Validation
## Create different types of data validation in Excel cells
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Decimal DataValidation
sheet.Range.get("B11").Text = "Input Number(3-6):";
let rangeNumber = sheet.Range.get("B12");
rangeNumber.DataValidation.CompareOperator =
  wasmModule.ValidationComparisonOperator.Between;
rangeNumber.DataValidation.Formula1 = "3";
rangeNumber.DataValidation.Formula2 = "6";
rangeNumber.DataValidation.AllowType =
  wasmModule.CellDataType.Decimal;
rangeNumber.DataValidation.ErrorMessage =
  "Please input correct number!";
rangeNumber.DataValidation.ShowError = true;
rangeNumber.Style.KnownColor =
  wasmModule.ExcelColors.Gray25Percent;

// Date DataValidation
sheet.Range.get("B14").Text = "Input Date:";
let rangeDate = sheet.Range.get("B15");
rangeDate.DataValidation.AllowType = wasmModule.CellDataType.Date;
rangeDate.DataValidation.CompareOperator =
  wasmModule.ValidationComparisonOperator.Between;
rangeDate.DataValidation.Formula1 = "1/1/1970";
rangeDate.DataValidation.Formula2 = "12/31/1970";
rangeDate.DataValidation.ErrorMessage = "Please input correct date!";
rangeDate.DataValidation.ShowError = true;
rangeDate.DataValidation.AlertStyle =
  wasmModule.AlertStyleType.Warning;
rangeDate.Style.KnownColor = wasmModule.ExcelColors.Gray25Percent;

// TextLength DataValidation
sheet.Range.get("B17").Text = "Input Text:";
let rangeTextLength = sheet.Range.get("B18");
rangeTextLength.DataValidation.AllowType =
  wasmModule.CellDataType.TextLength;
rangeTextLength.DataValidation.CompareOperator =
  wasmModule.ValidationComparisonOperator.LessOrEqual;
rangeTextLength.DataValidation.Formula1 = "5";
rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!";
rangeTextLength.DataValidation.ShowError = true;
rangeTextLength.DataValidation.AlertStyle =
  wasmModule.AlertStyleType.Stop;
rangeTextLength.Style.KnownColor =
  wasmModule.ExcelColors.Gray25Percent;

sheet.AutoFitColumn(2);
```

---

# Excel Group Management
## Expand and collapse groups in Excel worksheets
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Expand the grouped rows with ExpandCollapseFlags set to expand parent
sheet.Range.get("A16:G19").ExpandGroup(
  wasmModule.GroupByType.ByRows,
  wasmModule.ExpandCollapseFlags.ExpandParent
);

// Collapse the grouped rows
sheet.Range.get("A10:G12").CollapseGroup(wasmModule.GroupByType.ByRows);
```

---

# Excel Find and Replace Data
## Find specific text in Excel and replace it with new text, then highlight the cells
```javascript
// Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

// Find the "Area" string
let ranges = worksheet.FindAllString("Area", false, false);

// Traverse the found ranges
for (let range of ranges) {
  // Replace it with "Area Code"
  range.Text = "Area Code";
  // Highlight the color
  range.Style.Color = wasmModule.Color.get_Yellow();
}
```

---

# spire.xls javascript data
## find data in specific range
```javascript
// Specify a range in the worksheet
let range = sheet.Range.get({
  row: 1,
  column: 1,
  lastRow: 12,
  lastColumn: 8,
});

// Find text in the range
let textRanges = range.FindAllString("E-iceblue", false, false);

// Find numbers in the range
let numberRanges = range.FindAllNumber(100, true);
```

---

# spire.xls javascript find
## find string and number in excel cells
```javascript
// Find cells with the input string
let textRanges = sheet.FindAllString("E-iceblue", false, false);

// Create a string builder
let builder = [];

// Append the address of found cells in builder
if (textRanges.length !== 0) {
  for (let range of textRanges) {
    let address = range.RangeAddress;
    builder.push("The address of found text cell is: " + address);
  }
} else {
  builder.push("No cells that contain the text");
}

// Find cells with the input integer or double
let numberRanges = sheet.FindAllNumber(100, true);

// Append the address of found cells in builder
if (numberRanges.length !== 0) {
  for (let range of numberRanges) {
    let address = range.RangeAddress;
    builder.push("The address of found number cell is: " + address);
  }
} else {
  builder.push("No cells that contain the number");
}
```

---

# Excel Table Formatting
## Format table styles and settings in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Add Default Style to the table
sheet.ListObjects.get(0).BuiltInTableStyle =
  wasmModule.TableBuiltInStyles.TableStyleMedium9;

// Show Total
sheet.ListObjects.get(0).DisplayTotalRow = true;

// Set calculation type
sheet.ListObjects.get(0).Columns.get(0).TotalsRowLabel = "Total";
sheet.ListObjects.get(0).Columns.get(1).TotalsCalculation =
  wasmModule.ExcelTotalsCalculation.None;
sheet.ListObjects.get(0).Columns.get(2).TotalsCalculation =
  wasmModule.ExcelTotalsCalculation.None;
sheet.ListObjects.get(0).Columns.get(3).TotalsCalculation =
  wasmModule.ExcelTotalsCalculation.Sum;
sheet.ListObjects.get(0).Columns.get(4).TotalsCalculation =
  wasmModule.ExcelTotalsCalculation.Sum;

// Show table style row stripes and column stripes
sheet.ListObjects.get(0).ShowTableStyleRowStripes = true;
sheet.ListObjects.get(0).ShowTableStyleColumnStripes = true;
```

---

# spire.xls javascript controls
## insert various controls in Excel worksheet
```javascript
//Add a textbox
let textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100);
textbox.Text = "Hello World";

//Add a checkbox
let cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100);
cb.CheckState = wasmModule.CheckState.Checked;
cb.Text = "Check Box 1";

//Add a RadioButton
let rb = ws.RadioButtons.Add({
  row: 13,
  column: 2,
  height: 15,
  width: 100,
});
rb.Text = "Option 1";

// Add a combox
let cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100);
cbx.ListFillRange = ws.Range.get("A36:A42");
```

---

# spire.xls javascript html
## insert HTML string into Excel cell
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

const htmlCode = "<div>first line<br>second line<br>third line</div>";

// Insert HTML string into the cell
let range = sheet.Range.get("A1");
range.HtmlString = htmlCode;

// Save the workbook to the specified path
workbook.SaveToFile({
  fileName: "InsertHtmlStringIntoCell.xlsx",
  version: wasmModule.ExcelVersion.Version2013,
});
```

---

# Excel Text Replacement and Highlighting
## Replace specific text in Excel cells and highlight them with color
```javascript
// Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

let ranges = worksheet.FindAllString("Total", true, true);
for (let i = 0; i < ranges.Count; i++) {
  let range = ranges.get(i);
  // Reset the text, in other words, replace the text
  range.Text = "Sum";
  // Set the color
  range.Style.Color = wasmModule.Color.get_Yellow();
}
```

---

# Excel Data Retrieval and Extraction
## Extract rows containing "teacher" from one Excel file to a new Excel file
```javascript
// Create a new workbook instance and get the first worksheet.
let newBook = wasmModule.Workbook.Create();
let newSheet = newBook.Worksheets.get(0);

// Get the first worksheet from the existing workbook
let sheet = workbook.Worksheets.get(0);

// Retrieve data and extract it to the first worksheet of the new excel workbook
let i = 1;
let columnCount = sheet.Columns.Count;
let cells = sheet.Columns.get(0).Cells;
for (let j = 0; j < cells.Count; j++) {
  let range = cells.get(j);
  if (range.Text === "teacher") {
    let sourceRange = sheet.Range.get({
      row: range.Row,
      column: 1,
      lastRow: range.Row,
      lastColumn: columnCount,
    });
    let destRange = newSheet.Range.get({
      row: i,
      column: 1,
      lastRow: i,
      lastColumn: columnCount,
    });
    sheet.Copy({
      sourceRange: sourceRange,
      destRange: destRange,
      copyStyle: true,
    });
    i += 1;
  }
}
```

---

# spire.xls javascript data processing
## split Excel data into multiple columns
```javascript
// Split data into separate columns by the delimited characters - space.
let splitText = null;
let text = null;
let i = 1;
while (i < sheet.LastRow) {
  text = sheet.Range.get({ row: i + 1, column: 1 }).Text;
  splitText = text.split(" ");
  let j = 0;
  while (j < splitText.length) {
    sheet.Range.get({ row: i + 1, column: j + 2 }).Text = splitText[j];
    j += 1;
  }
  i += 1;
}
```

---

# Excel Subtotal Functionality
## Add subtotal to Excel data range
```javascript
// Select data range
let range = sheet.Range.get("A1:B18");

// Subtotal selected data
sheet.Subtotal({
  range: range,
  groupByIndex: 0,
  totalFields: [1],
  subtotalType: wasmModule.SubtotalTypes.Sum,
  replace: true,
  addPageBreak: false,
  addsummaryBelowData: true,
});
```

---

# Spire.XLS JavaScript Rich Text
## Write rich text with different font styles to Excel cells
```javascript
// Create an underlined font
let fontUnderline = workbook.CreateFont();
fontUnderline.Underline = wasmModule.FontUnderlineType.Single;

// Create an italic font
let fontItalic = workbook.CreateFont();
fontItalic.IsItalic = true;

// Create a green-colored font
let fontColor = workbook.CreateFont();
fontColor.KnownColor = wasmModule.ExcelColors.Green;

// Get the rich text object for cell B11
let richText = sheet.Range.get("B11").RichText;
richText.Text = "Bold and underlined and italic and colored text.";

// Apply the bold font to the range from character 0 to 3 (inclusive)
richText.SetFont(0, 3, fontBold);

// Apply the underline font to the range from character 9 to 18 (inclusive)
richText.SetFont(9, 18, fontUnderline);

// Apply the italic font to the range from character 24 to 29 (inclusive)
richText.SetFont(24, 29, fontItalic);

// Apply the green color font to the range from character 35 to 41 (inclusive)
richText.SetFont(35, 41, fontColor);
```

---

# Access Excel Cells
## Different methods to access cells in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Access cell by its name
let range1 = sheet.Range.get("A1");

// Access cell by index of row and column
let range2 = sheet.Range.get({ row: 2, column: 1 });

// Access cell in cell collection
let range3 = sheet.Cells.get(2);
```

---

# spire.xls javascript cell formatting
## apply multiple fonts in single cell
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Create a font object in workbook, setting the font color, size and type
let font1 = workbook.CreateFont();
font1.KnownColor = wasmModule.ExcelColors.LightBlue;
font1.IsBold = true;
font1.Size = 10;

// Create another font object specifying its properties
let font2 = workbook.CreateFont();
font2.KnownColor = wasmModule.ExcelColors.Red;
font2.IsBold = true;
font2.IsItalic = true;
font2.FontName = "Times New Roman";
font2.Size = 11;

// Write a RichText string to the cell 'H5', and set the font for it
let richText = sheet.Range.get("H5").RichText;
richText.Text = "This document was created with Spire.XLS for .NET.";
richText.SetFont(0, 29, font1);
richText.SetFont(31, 48, font2);
```

---

# spire.xls javascript autofit
## auto-fit column width and row height based on cell value
```javascript
//Set value for B8
let cell = worksheet.Range.get("B8");
cell.Text = "Welcome to Spire.XLS!";

//Set the cell style
let style = cell.Style;
style.Font.Size = 16;
style.Font.IsBold = true;

//Autofit column width and row height based on cell value
cell.AutoFitColumns();
cell.AutoFitRows();
```

---

# spire.xls javascript cell style
## get and find cells with same style name
```javascript
// Get the cell style name
let styleName = sheet.Range.get("A1").CellStyleName;

let ranges = sheet.AllocatedRange;
for (let cc of ranges.Cells) {
  // Find the cells which have the same style name
  if (cc.CellStyleName === styleName) {
    // Set value
    cc.Value = "Same style";
  }
}
```

---

# Excel Text to Number Conversion
## Convert text string format to number format in Excel cells
```javascript
// Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

// Convert text string format to number format
worksheet.Range.get("D2:D8").ConvertToNumber();
```

---

# spire.xls javascript format
## copy cell format from one column to another
```javascript
//Copy the cell format from column 2 and apply to cells of column 5
let count = sheet.Rows.Count;
for (let i = 1; i <= count; i++) {
  sheet.Range.get(`E${i}`).Style = sheet.Range.get(`B${i}`).Style;
}
```

---

# Count Cells in Excel Worksheet
## Get the total number of cells in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the number of cells
let cellCount = sheet.Cells.Count;
```

---

# Excel Cell Cutting Operation
## Cut cells from one position to another in Excel
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

let Ori = sheet.Range.get("A1:C5");
let Dest = sheet.Range.get("A26:C30");

// Copy the range to other position
sheet.Copy({
  sourceRange: Ori,
  destRange: Dest,
  copyStyle: true,
  updateReference: true,
  ignoreSize: true,
});

// Remove all content in original cells
for (let cr of Ori.Cells) {
  cr.ClearAll();
}
```

---

# Excel Merged Cells Detection
## Detect and unmerge merged cells in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the merged cell ranges in the first worksheet and put them into a CellRange array.
let range = sheet.MergedCells;

// Traverse through the array and unmerge the merged cells.
for (let cell of range) {
  cell.UnMerge();
}
```

---

# Spire.XLS JavaScript Duplicate Cell Range
## Duplicate a cell range in Excel with formatting
```javascript
// Copy data from source range to destination range and maintain the format
sheet.Copy({
  sourceRange: sheet.Range.get("A6:F6"),
  destRange: sheet.Range.get("A16:F16"),
  copyStyle: true,
});
```

---

# Excel Cell Clearing
## Methods to empty or clear cell contents in Excel
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Set the value as null to remove the original content from the Excel Cell
sheet.Range.get("C6").Value = "";

// Clear the contents to remove the original content from the Excel Cell
sheet.Range.get("B6").ClearContents();

// Remove the contents with format from the Excel cell
sheet.Range.get("D6").ClearAll();
```

---

# Excel Cell Color Filter
## Filter cells in Excel by cell color
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Create an auto filter in the sheet and specify the range to be filtered
sheet.AutoFilters.Range = sheet.Range.get("G1:G19");

// Get the column to be filtered
let filtercolumn = sheet.AutoFilters.get(0);

// Add a color filter to filter the column based on cell color
sheet.AutoFilters.AddFillColorFilter({
  filterColumnIndex: filtercolumn,
  color: wasmModule.Color.get_Red(),
});

// Filter the data
sheet.AutoFilters.Filter();
```

---

# Finding Cells with Style Name
## This code demonstrates how to find cells with a specific style name in an Excel worksheet and mark them
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the cell style name
let styleName = sheet.Range.get("A1").CellStyleName;

let ranges = sheet.AllocatedRange;
for (let cc of ranges.Cells) {
  // Find the cells which have the same style name
  if (cc.CellStyleName == styleName) {
    // Set value
    cc.Value = "Same style";
  }
}
```

---

# spire.xls javascript find formula cells
## find cells containing specific formula in excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Create a string builder
let builder = [];

// Find the cells that contain formula "=SUM(A5,A6)"
let ranges = sheet.FindAll(
  "=SUM(A5,A6)",
  wasmModule.FindType.Formula,
  wasmModule.ExcelFindOptions.None
);

// Append the address of found cells to builder
if (ranges.Count != 0) {
  for (let range of ranges) {
    let address = range.RangeAddress;
    builder.push(`The address of found cell is: ${address}`);
  }
} else {
  builder.push("No cell contain the formula");
}

// Combine all the found data into a single string
let content = builder.join("\n");
```

---

# JavaScript Excel Cell Address
## Get cell address and related information from Excel range
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Create a string builder
let builder = [];

// Get a cell range
let range = sheet.Range.get("A1:B5");

// Get address of range
let address = range.RangeAddressLocal;
builder.push(`Address of range: ${address}`);

// Get the cell count of range
let count = range.CellsCount;
builder.push(`Cell count of range: ${count}`);

// Get the address of the entire column of range
let entireColAddress = range.EntireColumn.RangeAddressLocal;
builder.push(
  `Address of entire column of the range: ${entireColAddress}`
);

// Get the address of the entire row of range
let entireRowAddress = range.EntireRow.RangeAddressLocal;
builder.push(`Address of entire row of the range ${entireRowAddress}`);

// Combine all the found data into a single string
let content = builder.join("\n");

// Clean up resources
workbook.Dispose();
```

---

# spire.xls javascript get cell data type
## get and display cell data types in excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the cell types of the cells in range "H2:H7"
for (let range of sheet.Range.get("H2:H7").Cells) {
  // Get cell type
  let cellType = sheet.GetCellType(range.Row, range.Column, false);
  
  // Write the cell type to the adjacent cell
  sheet.get({ row: range.Row, column: range.Column + 1 }).Text = cellType.toString();
  
  // Style the cell with red bold text
  sheet.get({
    row: range.Row,
    column: range.Column + 1,
  }).Style.Font.Color = wasmModule.Color.get_Red();
  sheet.get({
    row: range.Row,
    column: range.Column + 1,
  }).Style.Font.IsBold = true;
}
```

---

# Get Cell Displayed Text
## Extract cell value and displayed text in worksheet
```javascript
// Set value for B8
let cell = worksheet.Range.get("B8");
cell.NumberValue = 0.012345;

// Set the cell style
let style = cell.Style;
style.NumberFormat = "0.00";

// Get the cell value
let cellValue = cell.Value;

// Get the displayed text of the cell
let displayedText = cell.DisplayedText;
```

---

# Spire.XLS JavaScript Get Cell Value
## Retrieve cell value by cell name in Excel
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Create a string builder
let builder = [];

// Specify a cell by its name
let cell = sheet.Range.get("A2");

builder.push(`The value of cell A2 is: ${cell.Value}`);

// Combine all the found data into a single string
let content = builder.join("\n");
```

---

# Excel Range Intersection
## Get intersection of two ranges in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the intersection of two ranges
let range = sheet.Range.get("A2:D7").Intersect(
  sheet.Range.get("B2:E8")
);
```

---

# Hide Cell Content in Excel
## Hide cell content by setting number format
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Hide the area by setting the number format as ";;;"
sheet.Range.get("C5:D6").NumberFormat = ";;;";
```

---

# spire.xls javascript cells
## merge cells in Excel
```javascript
// Merge the seventh column in Excel file
workbook.Worksheets.get(0).Columns.get(6).Merge();

// Merge the particular range in Excel file
workbook.Worksheets.get(0).Range.get("A14:D14").Merge();
```

---

# Excel Cell Fill Pattern
## Set cell color and fill pattern in Excel using Spire.XLS for JavaScript
```javascript
// Set cell color
worksheet.Range.get("B7:F7").Style.Color =
  wasmModule.Color.get_Yellow();

// Set cell fill pattern
worksheet.Range.get("B8:F8").Style.FillPattern =
  wasmModule.ExcelPatternType.Percent125Gray;
```

---

# spire.xls javascript formatting
## set DB Num formatting for Excel cells
```javascript
// Get the cell range
let range = sheet.Range.get("A1:A3");

// Set the DB num format
range.NumberFormat = "[DBNum2][$-804]General";

// Auto fit columns
range.AutoFitColumns();
```

---

# Excel Cell Text Shrinking
## Shrink text to fit in a cell in an Excel file
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the cell range to shrink text
let cell = sheet.Range.get("B13:C13");

// Enable ShrinkToFit
let style = cell.Style;
style.ShrinkToFit = true;
```

---

# Traverse Cells Value
## Extract and display values from all cells in an Excel worksheet
```javascript
// Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

// Create a string builder
let builder = [];

// Get the cell range collection
let cellRangeCollection = worksheet.Cells;

builder.push("Values of the first sheet:");

// Traverse cells value
for (let cellRange of cellRangeCollection) {
  // Set string format for displaying
  let result = `Cell: ${cellRange.RangeAddress}   Value: ${cellRange.Value}`;

  // Add result string to content array
  builder.push(result);
}

// Combine all the found data into a single string
let content = builder.join("\n");
```

---

# Excel Cells Ungrouping
## Ungroup specific rows in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Ungroup the rows 10 to 12
sheet.UngroupByRows(10, 12);

// Ungroup the rows 16 to 19
sheet.UngroupByRows(16, 19);
```

---

# Excel Cell Unmerging
## Unmerge specific cells in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Unmerge the cells
sheet.Range.get("F2").UnMerge();
sheet.Range.get("F7").UnMerge();
```

---

# spire.xls javascript cells
## use explicit line breaks in worksheet cells
```javascript
// Specify a cell range
let c5 = worksheet.Range.get("C5");

// Set the cell width for specified range
worksheet.SetColumnWidth(c5.Column, 70);

// Put the string value with explicit line breaks
c5.Value =
  "Spire.XLS for JavaScript is a professional Excel API\n that can be used to create, read, \nwrite, convert and print Excel files";

// Set Text wrap
c5.IsWrapText = true;
```

---

# Excel Cell Text Wrapping
## Wrap or unwrap text in Excel cells
```javascript
// Wrap the excel text
sheet.Range.get("C1").Text =
  "e-iceblue is in facebook and welcome to like us";
sheet.Range.get("C1").Style.WrapText = true;
sheet.Range.get("D1").Text =
  "e-iceblue is in twitter and welcome to follow us";
sheet.Range.get("D1").Style.WrapText = true;

// Unwrap the excel text
sheet.Range.get("C2").Text =
  "http://www.facebook.com/pages/e-iceblue/139657096082266";
sheet.Range.get("C2").Style.WrapText = false;
sheet.Range.get("D2").Text = "https://twitter.com/eiceblue";
sheet.Range.get("D2").Style.WrapText = false;
```

---

# spire.xls javascript autofit column
## autofit column in specific range
```javascript
// Autofit the Column of the worksheet
sheet.AutoFitColumn(2, 2, 5);
```

---

# Excel Row AutoFit in Range
## AutoFit a specific row within a specified column range in Excel worksheet
```javascript
// Get the first worksheet
const sheet = workbook.Worksheets.get(0);

// Autofit the second row of the worksheet within columns 1 to 2
sheet.AutoFitRow({rowIndex:2, firstColumn:1, lastColumn:2, options:false});
```

---

# spire.xls javascript autofit check
## check if excel row or column is auto fit
```javascript
// Gets whether the cell has an adaptive row height set
const isRowAutofit = workbook.Worksheets.get(0).GetRowIsAutoFit(2);
if (isRowAutofit) {
  result.push("The second row is auto fit row height.");
} else {
  result.push("The second row is not auto fit row height.");
}

// Gets whether the cell has an adaptive column width set
const isColAutofit = workbook.Worksheets.get(0).GetColumnIsAutoFit(2);
if (isColAutofit) {
  result.push("The second column is auto fit column width.");
} else {
  result.push("The second column is not auto fit column width.");
}
```

---

# Check Hidden Rows or Columns in Excel
## Determine if specific rows or columns are hidden in an Excel worksheet
```javascript
// Get the first worksheet in the workbook
let sheet = workbook.Worksheets.get(0);
let result = [];
let rowIndex = 2;
let columnIndex = 2;

// Check if row is hidden
let rowIsHide = sheet.GetRowIsHide(rowIndex);
if (rowIsHide) {
  result.push("The second row is hidden.");
} else {
  result.push("The second row is not hidden.");
}

// Check if column is hidden
let columnIsHide = sheet.GetColumnIsHide(columnIndex);
if (columnIsHide) {
  result.push("The second column is hidden.");
} else {
  result.push("The second column is not hidden.");
}
```

---

# Spire.XLS JavaScript Column Copying
## Copy columns within and between Excel worksheets
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get the first and second sheet
let sheet1 = workbook.Worksheets.get(0);
let sheet2 = workbook.Worksheets.get(1);

// Copy the first column to the third column in the same sheet
sheet1.Copy({sourceRange:sheet1.Columns.get(0), destRange:sheet1.Columns.get(2), copyStyle:true, updateReference:true, ignoreSize:true});

// Copy the first column to the second column in the different sheet
sheet1.Copy({sourceRange:sheet1.Columns.get(0), destRange:sheet2.Columns.get(1), copyStyle:true, updateReference:true, ignoreSize:true});
```

---

# spire.xls javascript copy rows
## copy rows within and between excel sheets
```javascript
//Get the first and second sheet
let sheet1 = workbook.Worksheets.get(1);
let sheet2 = workbook.Worksheets.get(0);

//Copy the first row to the third row in the same sheet
sheet1.Copy({sourceRange:sheet1.Rows.get(0), destRange:sheet1.Rows.get(2), copyStyle:true, updateReference:true, ignoreSize:true});

//Copy the first row to the second row in the different sheet
sheet1.Copy({sourceRange:sheet1.Rows.get(0), destRange:sheet2.Rows.get(1), copyStyle:true, updateReference:true, ignoreSize:true});
```

---

# spire.xls javascript copy column and row
## copy single column and row to specified destination ranges in excel worksheet
```javascript
//Load the document
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: excelFileName});

//Get the first sheet
const sheet1 = workbook.Worksheets.get(0);

// Specify a destination range to copy one column
const columnCells = sheet1.Range.get("G1:G19");

// Copy the second column to destination range
sheet1.Columns.get(1).Copy({destRange:columnCells});

// Specify a destination range to copy one row
const rowCells = sheet1.Range.get("A21:E21");

// Copy the first row to destination range
sheet1.Rows.get(0).Copy({destRange:rowCells});
```

---

# Excel Range Copy with Options
## Copy cell range with style preservation and reference updating
```javascript
//Get the first worksheet
let sheet1 = workbook.Worksheets.get(0);

//Add a new worksheet as destination sheet
let destinationSheet = workbook.Worksheets.Add("DestSheet");

//Specify a copy range of original sheet
let cellRange = sheet1.Range.get("B2:D4");

//Copy the specified range to added worksheet and keep original styles and update reference
workbook.Worksheets.get(0).Copy({sourceRange:cellRange, worksheet:destinationSheet, destRow:2, destColumn:1,copyStyle:true, updateReference:true});
```

---

# spire.xls javascript delete blank rows and columns
## delete blank rows and columns in Excel worksheet
```javascript
//Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

//Delete blank rows from the worksheet.
for(let i=sheet.Rows.Count-1; i>=0; i--) {
  if(sheet.Rows.get(i).IsBlank) {
    sheet.DeleteRow(i+1);
  }
}

//Delete blank columns from the worksheet.
for(let j=sheet.Columns.Count-1; j>=0; j--) {
  if(sheet.Columns.get(j).IsBlank) {
    sheet.DeleteColumn(j+1);
  }
}
```

---

# Spire.XLS JavaScript Rows and Columns
## Delete multiple rows and columns from Excel worksheet
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Delete 4 rows from the fifth row
sheet.DeleteRow({index:5, count:4});

//Delete 2 columns from the second column
sheet.DeleteColumn({index:2, count:2});
```

---

# Spire.XLS JavaScript Get Default Row and Column Count
## Get the default row and column count of a worksheet
```javascript
//Create a workbook
let workbook = spirexls.Workbook.Create();

//Clear all worksheets
workbook.Worksheets.Clear();

//Create a new worksheet
let sheet = workbook.CreateEmptySheet();
let sb = [];

//Get row and column count
let rowCount = sheet.Rows.Count;
let columnCount = sheet.Columns.Count;
sb.push(`The default row count is :${rowCount}`);
sb.push(`The default column count is :${columnCount}`);
```

---

# Spire.XLS JavaScript Grouping
## Group rows and columns in an Excel worksheet
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Grouping rows
sheet.GroupByRows(1, 5, false);

//Grouping columns
sheet.GroupByColumns(1, 3, false);
```

---

# Hide or Show Row Column Headers in Excel
## This code demonstrates how to hide or show row and column headers in an Excel worksheet using Spire.XLS for JavaScript.
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);

// Hide the headers of rows and columns
sheet.RowColumnHeadersVisible = false;
```

---

# Hide Rows and Columns in Excel
## This code demonstrates how to hide specific rows and columns in an Excel worksheet using JavaScript
```javascript
// Load the document
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: excelFileName});

// Get the first sheet
let worksheet = workbook.Worksheets.get(0);

// Hiding the column of the worksheet
worksheet.HideColumn(2);

// Hiding the row of the worksheet
worksheet.HideRow(4);
```

---

# Spire.XLS JavaScript Rows and Columns
## Insert rows and columns in Excel worksheet
```javascript
//Get the first sheet
let worksheet = workbook.Worksheets.get(0);

//Inserting a row into the worksheet
worksheet.InsertRow(2);

//Inserting a column into the worksheet
worksheet.InsertColumn(2);

//Inserting multiple rows into the worksheet
worksheet.InsertRow({rowIndex:5, rowCount:2});

//Inserting multiple columns into the worksheet
worksheet.InsertColumn({columnIndex:5, columnCount:2});
```

---

# Remove Excel Row Based on Keyword
## Delete a row from an Excel worksheet that contains a specific keyword
```javascript
//Load the document
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: excelFileName});

//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Find the string
let cr = sheet.FindString("Address", false, false);

//Delete the row which includes the string
sheet.DeleteRow(cr.Row);

//Dispose
workbook.Dispose();
```

---

# Spire.XLS JavaScript Column Width
## Set column width in pixels
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Set the width of the third column to 400 pixels
sheet.SetColumnWidthInPixels(3, 400);
```

---

# Set Default Column Width in Excel
## Set the default width for columns in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Set default column width
sheet.DefaultColumnWidth = 25;

// Save result file
workbook.SaveToFile({fileName: outputFileName, version:wasmModule.ExcelVersion.Version2010});

// Dispose
workbook.Dispose();
```

---

# Excel Default Row and Column Style
## Set default style for rows and columns in Excel spreadsheet
```javascript
//Create the document
let workbook = spirexls.Workbook.Create();

//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Create a cell style and set the color
let style = workbook.Styles.Add("Mystyle");
style.Color = spirexls.Color.get_Yellow();

//Set the default style for the first row and column
sheet.SetDefaultRowStyle(1, style);
sheet.SetDefaultColumnStyle(1, style);
```

---

# spire.xls javascript row height
## set default row height for excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Set default row height
sheet.DefaultRowHeight = 30;
```

---

# Excel Row Height and Column Width Setting
## Set row height and column width in Excel worksheet
```javascript
//Get the first sheet
let worksheet = workbook.Worksheets.get(0);

// Setting the width to 30
worksheet.SetColumnWidth(4, 30);

// Setting the height to 30
worksheet.SetRowHeight(4, 30)
```

---

# Spire.XLS JavaScript Column Direction
## Set summary column direction in Excel
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Group Columns
sheet.GroupByColumns(1, 4, true);

//Set summary columns to right of details
sheet.PageSetup.IsSummaryRowBelow = true;
```

---

# Excel Summary Row Direction
## Set summary row position above or below detail rows
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Group rows
sheet.GroupByRows(1, 4, true);

//Set summary rows above details
sheet.PageSetup.IsSummaryRowBelow = false;
```

---

# Unhide Rows and Columns in Excel
## This code demonstrates how to unhide specific rows and columns in an Excel worksheet
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Unhide the row
sheet.ShowRow(15);

//Unhide the column
sheet.ShowColumn(4);
```

---

# spire.xls javascript image alignment
## align picture within excel cell
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Get the first worksheet in the workbook
let sheet = workbook.Worksheets.get(0);
// Set the text in cell A1
sheet.Range.get("A1").Text = "Align Picture Within A Cell:";
// Set the vertical alignment of cell A1 to top
sheet.Range.get("A1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Top;
// Insert an image at the specific cell (1, 1)
let picture = sheet.Pictures.Add({topRow:1, leftColumn:1, fileName:inputFileName});
// Adjust the column width and row height so that the cell can contain the picture
sheet.Columns.get(0).ColumnWidth = 40;
sheet.Rows.get(0).RowHeight =200;
// Set the horizontal offset of the image within the cell to 100
picture.LeftColumnOffset = 100;
// Set the vertical offset of the image within the cell to 25
picture.TopRowOffset = 25;
```

---

# Spire.XLS JavaScript Image Compression
## Compress pictures in Excel workbook
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
// Compress the picture quality for all pictures in all worksheets
for(let sheet of workbook.Worksheets) {
    for(let picture of sheet.Pictures) {
      // Set the compression level to 50 (50% of original quality)  
      picture.Compress(50);
    }
}
// Save the modified workbook to the specified file
workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Picture Copy
## Copy picture from one worksheet to another in Excel
```javascript
// Create a new workbook 
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
//Get the first worksheet
let sheet1 = workbook.Worksheets.get(0);
//Add a new worksheet as destination sheet
let destinationSheet = workbook.Worksheets.Add("DestSheet");
//Get the first picture from the first worksheet
let sourcePicture = sheet1.Pictures.get(0);
//Get the image
let image = sourcePicture.Picture;
//Add the image into the added worksheet 
destinationSheet.Pictures.Add({topRow:2, leftColumn:2, stream:image});
```

---

# Get Cropped Position of Picture
## Extract the cropped position (left, top, width, height) of a picture in an Excel worksheet
```javascript
// Get the first worksheet
let sheet1 = workbook.Worksheets.get(0);

// Get the image from the first sheet
let picture = sheet1.Pictures.get(0);

// Get the cropped position
let left = picture.Left;
let top = picture.Top;
let width = picture.Width;
let height = picture.Height;

// Set string format for displaying
let displayString = `Crop position: Left ${left}\r\nCrop position: Top ${top}\r\nCrop position: Width ${width}\r\nCrop position: Height ${height}`;
```

---

# spire.xls javascript images
## Delete all images from Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Delete all images from the worksheet
for (let i = sheet.Pictures.Count - 1; i >=0; i--) {
    sheet.Pictures.get(i).Remove();
}
```

---

# Excel Background Image Insertion
## Inserts a background image to an Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile(inputFileName2);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Open an image
let bm = wasmModule.Stream.CreateByFile(inputFileName1);

// Set the image to be background image of the worksheet
sheet.PageSetup.BackgoundImage = bm;
```

---

# Excel Image Positioning
## Locate and adjust image position in Excel worksheet
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);
// Get the first picture from the sheet
let pic = sheet.Pictures.get(0);
// Set the horizontal offset of the picture within the cell to 300
pic.LeftColumnOffset = 300;
// Set the vertical offset of the picture within the cell to 300
pic.TopRowOffset = 300;
```

---

# Picture Offset in Excel
## Set left column and top row offset for a picture in Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Insert a picture
let pic = sheet.Pictures.Add({topRow:2, leftColumn:2, fileName:inputFileName});

// Set left offset and top offset from the current range
pic.LeftColumnOffset = 200;
pic.TopRowOffset = 100;
```

---

# Excel Picture Reference Range
## Setting a reference range for a picture in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);
// Set values in cells A1 and B3.
sheet.Range.get("A1").Value = "Spire.XLS";
sheet.Range.get("B3").Value = "E-iceblue";

// Get the first picture in worksheet
let picture = sheet.Pictures.get(0);

// Set the reference range of the picture to A1:B3
picture.RefRange = "A1:B3";
```

---

# spire.xls javascript images
## read image from worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);
//Get the first image 
let pic = sheet.Pictures.get(0);
const outputFileName = 'ReadImages-out.png';

pic.Picture.Save(outputFileName);

workbook.Dispose();
```

---

# Spire.XLS JavaScript Picture Border
## Remove picture border in Excel worksheet
```javascript
// Get the first picture from the first worksheet
let picture = sheet1.Pictures.get(0);

// Remove the picture border
// Method-1:
picture.Line.Visible = false;

// Method-2:
// picture.Line.Weight = 0
```

---

# Excel Image Size and Position Reset
## Reset the size and position of an image in an Excel worksheet
```javascript
// Add a picture to the first worksheet
let picture = sheet.Pictures.Add({topRow:1, leftColumn:1, fileName:inputFileName});

// Set the size for the picture
picture.Width = 200;
picture.Height = 200;

// Set the position for the picture
picture.Left = 200;
picture.Top = 100;
```

---

# spire.xls javascript chart
## set image offset of chart
```javascript
// Add chart1 and background image to sheet1 as comparison.
let chart1 = sheet1.Charts.Add({chartType:wasmModule.ExcelChartType.ColumnClustered});
chart1.DataRange = sheet.Range.get("D1:E8");
chart1.SeriesDataFromRange = false;

// Chart Position.
chart1.LeftColumn = 1;
chart1.TopRow = 11;
chart1.RightColumn = 8;
chart1.BottomRow = 33;

let bm = wasmModule.Stream.CreateByFile(inputFileName2);
// Add picture as background.
chart1.ChartArea.Fill.CustomPicture({im:bm,name:"None"});
chart1.ChartArea.Fill.Tile = false;

// Set the image offset.
chart1.ChartArea.Fill.PicStretch.Left = 20;
chart1.ChartArea.Fill.PicStretch.Top = 20;
chart1.ChartArea.Fill.PicStretch.Right = 5;
chart1.ChartArea.Fill.PicStretch.Bottom = 5;
```

---

# spire.xls javascript images
## insert image into worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Add an image to the specific cell
sheet.Pictures.Add({topRow:14, leftColumn:5, fileName:inputFileName2});
```

---

# Excel Comment with Author
## Add a comment with author information to an Excel cell
```javascript
//Get the range that will add comment
let range = sheet.Range.get("C1");

//Set the author and comment content
let author = "E-iceblue";
let text = "This is demo to show how to add a comment with editable Author property.";

//Add comment to the range and set properties
let comment = range.AddComment();
comment.Width = 200;
comment.Visible = true;
comment.Text = author + ":\n" + text;

//Set the font of the author
let font = workbook.CreateFont();
font.FontName = "Arial";
font.KnownColor = wasmModule.ExcelColors.Black;
font.IsBold = true;
comment.RichText.SetFont(0, author.length, font);
```

---

# spire.xls javascript comment
## add comment with picture to Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);
// Set value for the range
sheet.Range.get("C6").Text = "E-iceblue";
//Add the comment
let comment = sheet.Range.get("C6").AddComment();
// Fill the comment with a customized background picture
comment.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile(inputFileName), name:"None"});
comment.Visible = true;
```

---

# Excel Comment Editing
## Edit existing comment in an Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Get the first comment
let comment = sheet.Comments.get(0);

// Edit the comment
comment.Text = "This comment has been edited by Spire.XLS.";

// Save the modified workbook to the specified file
workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Comment Visibility Control
## Hide or show comments in Excel worksheet
```javascript
//Hide the second comment
sheet.Comments.get(1).IsVisible = false;

//Show the third comment
sheet.Comments.get(2).IsVisible = true;
```

---

# Excel Comment Reader
## Read comments from Excel cells
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
// Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

// Get the comment text
let commentText1 = sheet.Range.get("A1").Comment.Text;
let commentText2 = sheet.Range.get("A2").Comment.RichText.RtfText;

workbook.Dispose();
```

---

# Excel Comment Management
## Remove and modify Excel comments
```javascript
//Get all comments of the first sheet
let comments = workbook.Worksheets.get(0).Comments;
//Change the content of the first comment
comments.get(0).Text = "This comment has been changed.";
//Remove the second comment
comments.get(1).Remove();
```

---

# Excel Comment Fill Color
## Set fill color for Excel comment
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

//Create Excel font
let font = workbook.CreateFont();
font.FontName = "Arial";
font.Size = 11;
font.KnownColor = wasmModule.ExcelColors.Orange;

//Add the comment
let range = sheet.Range.get("A1");
let commentText = "This is a comment";
range.Comment.Text = commentText;
range.Comment.RichText.SetFont(0, commentText.length - 1, font);

//Set comment Color
range.Comment.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
range.Comment.Fill.ForeColor = wasmModule.Color.get_SkyBlue();
range.Comment.Visible = true;
```

---

# spire.xls javascript comment
## set comment text rotation
```javascript
// Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

//Create Excel font
let font = workbook.CreateFont();
font.FontName = "Arial";
font.Size = 11;
font.KnownColor = wasmModule.ExcelColors.Orange;

//Add the comment
let range = sheet.Range.get("E1");
let commentText = "This is a comment";
range.Comment.Text = commentText;
range.Comment.RichText.SetFont(0, commentText.length - 1, font);

// Set its vertical and horizontal alignment
range.Comment.VAlignment = wasmModule.CommentVAlignType.Center;
range.Comment.HAlignment = wasmModule.CommentHAlignType.Right;

//Set the comment text rotation
range.Comment.TextRotation = wasmModule.TextRotationType.LeftToRight;
```

---

# spire.xls javascript comments
## set position and alignment for comments
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

//Set two font styles which will be used in comments
let font1 = workbook.CreateFont();
font1.FontName = "Calibri";
font1.Color = wasmModule.Color.get_Firebrick();
font1.IsBold = true;
font1.Size = 12;
let font2 = workbook.CreateFont();
font2.FontName = "Calibri";
font2.Color = wasmModule.Color.get_Blue();
font2.Size = 12;
font2.IsBold = true;

//Add comment 1 and set its size, text, position and alignment
sheet.Range.get("G5").Text = "Spire.XLS";
let Comment1 = sheet.Range.get("G5").Comment;
Comment1.IsVisible = true;
Comment1.Height = 150;
Comment1.Width = 300;
Comment1.RichText.Text = "Spire.XLS for JavaScript:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ";
Comment1.RichText.SetFont(0, 19, font1);
Comment1.TextRotation = wasmModule.TextRotationType.LeftToRight;

//Set the position of Comment
Comment1.Top = 20;
Comment1.Left = 40;

//Set the alignment of text in Comment
Comment1.VAlignment = wasmModule.CommentVAlignType.Center;
Comment1.HAlignment = wasmModule.CommentHAlignType.Justified;

//Add comment2 and set its size, text, position and alignment for comparison
sheet.Range.get("D14").Text = "E-iceblue";
let Comment2 = sheet.Range.get("D14").Comment;
Comment2.IsVisible = true;
Comment2.Height = 150;
Comment2.Width = 300;
Comment2.RichText.Text = "About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.";
Comment2.TextRotation = wasmModule.TextRotationType.LeftToRight;
Comment2.RichText.SetFont(0, 16, font2);
//Set the position of Comment
Comment2.Top = 170;
Comment2.Left = 450;
//Set the alignment of text in Comment
Comment2.VAlignment = wasmModule.CommentVAlignType.Top;
Comment2.HAlignment = wasmModule.CommentHAlignType.Justified;
```

---

# spire.xls javascript comments
## write comments to excel cells
```javascript
//Creates font
let font = workbook.CreateFont();
font.FontName = "Arial";
font.Size = 11;
font.KnownColor = wasmModule.ExcelColors.Orange;
let fontBlue = workbook.CreateFont();
fontBlue.KnownColor = wasmModule.ExcelColors.LightBlue;
let fontGreen = workbook.CreateFont();
fontGreen.KnownColor = wasmModule.ExcelColors.LightGreen;

let range = sheet.Range.get("B11");
range.Text = "Regular comment";
range.Comment.Text = "Regular comment";
range.AutoFitColumns();

range = sheet.Range.get("B12");
range.Text = "Rich text comment";
range.RichText.SetFont(0, 16, font);
range.AutoFitColumns();
//Rich text comment
range.Comment.RichText.Text = "Rich text comment";
range.Comment.RichText.SetFont(0, 4, fontGreen);
range.Comment.RichText.SetFont(5, 9, fontBlue);
```

---

# Excel Chart Sheet to SVG Conversion
## Convert Excel chart sheet to SVG format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

//Get the chartsheet by name
let cs = workbook.GetChartSheetByName("Chart1");
const outputFileName = 'ChartSheetToSVG-out.svg';
let outStream = wasmModule.Stream.CreateByFile(outputFileName);
cs.ToSVGStream(outStream);
cs.Flush();
cs.Close();
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# CSV to Excel Conversion
## Convert CSV files to Excel format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing CSV document
workbook.LoadFromFile({ fileName: inputFileName, separator: ",", row: 1, column: 1 });
// Get the first worksheet.
let sheet = workbook.Worksheets.get(0);
// Ignore error options for the range D2:E19, treating numbers as text
sheet.Range.get("D2:E19").IgnoreErrorOptions = wasmModule.IgnoreErrorType.NumberAsText;
// Auto-fit columns in the allocated range of the worksheet
sheet.AllocatedRange.AutoFitColumns();

const outputFileName = 'CSVToExcel-out.xlsx';
// Save the modified workbook to the specified file
workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# CSV to PDF Conversion
## Convert CSV file to PDF format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing CSV document
workbook.LoadFromFile({ fileName: inputFileName, separator: ",", row: 1, column: 1 });

// Set the SheetFitToPage property as true
workbook.ConverterSetting.SheetFitToPage = true;

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Autofit a column if the characters in the column exceed column width
for (let i = 1; i < sheet.Columns.Count; i++) {
    sheet.AutoFitColumn(i);
}

// Save to PDF document
workbook.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Worksheet to PDF Conversion
## Convert each worksheet in an Excel workbook to a separate PDF file

```javascript
//Save each sheet to PDF
for (let i = 0; i < workbook.Worksheets.Count; i++) {
    let sheet = workbook.Worksheets.get(i);    
    const outputFileName = sheet.Name + '.pdf';
    sheet.SaveToPdf({fileName: outputFileName});
}
```

---

# Spire.XLS JavaScript Excel to PDF Conversion
## Fit to one page width when converting Excel to PDF
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

for (let i = 0; i < workbook.Worksheets.Count; i++) {
    let sheet = workbook.Worksheets.get(i);
    // Auto fit page height
    sheet.PageSetup.FitToPagesTall = 0;
    // Fit one page width
    sheet.PageSetup.FitToPagesWide = 1;
}

const outputFileName = 'FitWidthWhenConvertToPDF-out.pdf';
//Save to PDF document
workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PDF});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# HTML to Excel Conversion
## Convert HTML file to Excel format using Spire.XLS for JavaScript

```javascript
let inputFileName='HtmlToExcel.html';
await wasmModule.FetchFileToVFS(inputFileName, '', '');

// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing HTML document
workbook.LoadFromHtml({fileName: inputFileName});
    
const outputFileName = 'HtmlToExcel-out.xlsx';
// Save the modified workbook to the specified file
workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Load and Save ET and ETT Files
## This code demonstrates how to load an ET file and save it in ET format
```javascript
wasmModule = window.wasmModule;
if (wasmModule) {
  let inputFileName='LoadSaveEtAndETT.et';
  await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

  // Create a new workbook
  const workbook = wasmModule.Workbook.Create();
  // Load an existing Excel document
  workbook.LoadFromFile({fileName: inputFileName});      

  const outputFileName = 'LoadSaveEtAndETT-out.et';
  // Save to ET document
  workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.ET});
  // Dispose of the workbook object to release resources
  workbook.Dispose();
}
```

---

# Office Open XML to Excel Conversion
## Convert Office Open XML format to Excel format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
let fileStream = wasmModule.Stream.CreateByFile(inputFileName);
// Load an existing XML document
workbook.LoadFromXml({stream:fileStream});

const outputFileName = 'OfficeOpenXMLToExcel-out.xlsx';
// Save the modified workbook to the specified file using Excel 2013 format
workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Range to PDF Conversion
## Convert a selected range of cells from Excel to PDF format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

// Add a new sheet to workbook
workbook.Worksheets.Add("newsheet");
// Copy your area to new sheet.
workbook.Worksheets.get(0).Range.get("A9:E15").Copy({destRange:workbook.Worksheets.get(1).Range.get("A9:E15"), updateReference:false, copyStyles:true});
// Auto fit column width
workbook.Worksheets.get(1).Range.get("A9:E15").AutoFitColumns();

const outputFileName = 'SelectedRangeToPDF-out.pdf';
// Save worksheet to PDF
workbook.Worksheets.get(1).SaveToPdf({fileName:outputFileName});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls javascript conversion
## convert worksheet to image
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Convert the sheet to image
let image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);
```

---

# spire.xls javascript conversion
## convert specific worksheet cells to image
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

// Get the first sheet
let sheet = workbook.Worksheets.get(0);

const outputFileName ='SpecificCellsToImage-out.jpg';

// Specify Cell Ranges and Save to certain Image formats
sheet.ToImage(8, 1, 15, 5).Save(outputFileName);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Font Directory Specification
## Specify font directory when converting Excel to PDF
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
     
// Specify font directory
workbook.CustomFontFileDirectory = ["/Library/Fonts/"];

const outputFileName = 'SpecifyFontDirectory-out.pdf';
// Save to pdf file  
workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PDF});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel to CSV Conversion
## Convert Excel worksheet to CSV format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

// Get the first sheet
let sheet = workbook.Worksheets.get(0);

const outputFileName = 'ToCSV-out.csv';
// Convert to CSV file
sheet.SaveToFile({fileName:outputFileName, separator:",", encoding:wasmModule.Encoding.get_UTF8()});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel to CSV Conversion
## Convert Excel worksheet to CSV format with filtered values
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'ToCSVWithFilteredValue-out.csv';
// Save a worksheet to CSV
workbook.Worksheets.get(0).SaveToFile({fileName:outputFileName, separator:";", retainHiddenData:false});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel to HTML Conversion
## Convert Excel worksheet to HTML format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

// Get the first sheet
let sheet = workbook.Worksheets.get(0);

// Create HTML options for saving to HTML format
let options = wasmModule.HTMLOptions.Create();
// Embed images in the HTML file
options.ImageEmbedded = true;

// Save to HTML
sheet.SaveToHtml({fileName:outputFileName, saveOption:options});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Conversion
## Convert worksheet to HTML stream
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

// Get the first sheet
let sheet = workbook.Worksheets.get(0);

// Set the html options
let options = wasmModule.HTMLOptions.Create();
options.ImageEmbedded = true;

// Save sheet to html stream
let fileStream = wasmModule.Stream.CreateByFile(outputFileName);
sheet.SaveToHtml({stream:fileStream, saveOption:options});

// Dispose of the object to release resources
fileStream.Dispose();
workbook.Dispose();
```

---

# Spire.XLS JavaScript Conversion
## Convert worksheet to image without white space
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);
//Set the margin as 0 to remove the white space around the image
sheet.PageSetup.LeftMargin = 0;
sheet.PageSetup.BottomMargin = 0;
sheet.PageSetup.TopMargin = 0;
sheet.PageSetup.RightMargin = 0;
//convert to image
let image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);
```

---

# Excel to ODS Conversion
## Convert Excel file to ODS format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName='ToODS-out.ods';
// Save to ODS
workbook.SaveToFile({fileName:outputFileName,fileFormat:wasmModule.FileFormat.ODS});
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Conversion
## Convert Excel to OFD format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'ToOFD-out.ofd';
// Save to OFD
workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.OFD});
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Excel to Office Open XML Conversion
## Convert Excel workbook to Office Open XML format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Get the first sheet
let sheet = workbook.Worksheets.get(0);
// Set the text "Hello World" in cell A1 of the worksheet 
sheet.Range.get("A1").Text = "Hello World";
// Apply the color Gray25Percent to cell B1 using a known color
sheet.Range.get("B1").Style.KnownColor = wasmModule.ExcelColors.Gray25Percent;
// Apply the color Gold to cell C1 using a known color
sheet.Range.get("C1").Style.KnownColor = wasmModule.ExcelColors.Gold;

const outputFileName = 'ToOfficeOpenXML-out.xml';
// Save the workbook as an XML file
workbook.SaveAsXml({fileName:outputFileName});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel to PDF Conversion
## Convert Excel workbook to PDF format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'ToPDF-out.pdf';
// Save to PDF
workbook.SaveToFile({fileName: outputFileName , fileFormat: wasmModule.FileFormat.PDF});
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Excel to PDF/A-1B Conversion
## Convert Excel files to PDF/A-1B format using JavaScript

```javascript
// Load input Excel file
let inputFileName = 'ToPDFA1B.xlsx';
await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

// Create a workbook and load the Excel file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: inputFileName});

// Set PDF conformance level to PDF/A-1B
workbook.ConverterSetting.PdfConformanceLevel = wasmModule.PdfConformanceLevel.Pdf_A1B;

// Save as PDF
const outputFileName = 'ToPDFA1B-out.pdf';
workbook.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});
workbook.Dispose();
```

---

# Excel to PDF Conversion
## Simple conversion of Excel file to PDF format
```javascript
let inputFileName = 'ToPDF.xlsx';

// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'ToPDFSimply-out.pdf';
// Save to PDF
workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PDF});
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Excel to PDF Conversion with Page Size Change
## Convert Excel file to PDF while changing the page size to A3
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

for(let sheet of workbook.Worksheets) {
    // Change the page size
    sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperA3;
}

const outputFileName = 'ToPdfWithChangePageSize-out.pdf';
// Save to PDF
workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PDF});
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Excel to PostScript Conversion
## Convert Excel workbook to PostScript format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'ToPostScript-out.ps';
// Save to PostScript
workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PostScript});
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Excel to SVG Conversion
## Convert Excel worksheets to SVG format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

let i=0;
for(let worksheet of workbook.Worksheets) {
    const outputFileName = "sheet-"+i+".svg";
    // Create a FileStream to write the SVG content to a file
    let fs = wasmModule.Stream.CreateByFile(outputFileName);
    // Convert the worksheet to SVG and write it to the FileStream
    worksheet.ToSVGStream(fs,0, 0,0, 0);
    fs.Flush();
    fs.Dispose();
    i++;
}
// Dispose of the object to release resources
workbook.Dispose();
```

---

# Excel to Text Conversion
## Convert Excel worksheet to text format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
// Get the first worksheet in excel workbook
let sheet = workbook.Worksheets.get(0);
const outputFileName = 'ToText-out.txt';
// Save to text
sheet.SaveToFile({fileName:outputFileName, separator:" ", encoding:wasmModule.Encoding.get_UTF8()});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel to XPS Conversion
## Convert Excel file to XPS format using Spire.XLS for JavaScript
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'ToXPS-out.xps';
// Save to XPS
workbook.SaveToFile({fileName: outputFileName , fileFormat: wasmModule.FileFormat.XPS});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Workbook to HTML Conversion
## Convert Excel workbook to HTML format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({
  fileName: outputDirectoryName + inputFileName,
});

const outputFileName = "WorkbookToHTML-out.html";
// Save to HTML
workbook.SaveToHtml({
  fileName: outputDirectoryName + outputFileName,
});

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Spire.XLS JavaScript Conversion
## Convert XLS to XLSM format
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

const outputFileName = 'XLSToXLSM-out.xlsm';
// Save to XLSM
workbook.SaveToFile({ fileName: outputFileName, version: wasmModule.ExcelVersion.Version2007});
// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel AutoFilter Blank Data
## Filter blank cells in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Match the blank data
sheet.AutoFilters.MatchBlanks(0);

// Filter
sheet.AutoFilters.Filter();
```

---

# Excel AutoFilter Non-Blank Data
## Apply auto-filter to show only non-blank cells in Excel
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);
// Match the non blank data
sheet.AutoFilters.MatchNonBlanks(0);
// Filter 
sheet.AutoFilters.Filter();
```

---

# Excel Filter Creation
## Create filter for a specific cell range in Excel worksheet
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Create filter
sheet.AutoFilters.Range = sheet.Range.get("A1:J1");
```

---

# spire.xls javascript filter
## filter cells by string in Excel
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);

// Filter cells data which start with "South"
sheet.AutoFilters.Range = sheet.Range.get("D1:D19");
let filterColumn = sheet.AutoFilters.get(0);
let strCrt = "South*";
sheet.AutoFilters.CustomFilter({column:filterColumn, operatorType:wasmModule.FilterOperatorType.Equal, criteria:wasmModule.String.Create(strCrt)});
sheet.AutoFilters.Filter();
```

---

# spire.xls javascript data validation
## get settings of data validation from worksheet cells
```javascript
//Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

//Cell B4 has the Decimal Validation
let cell = worksheet.Range.get("B4");

//Get the validation of this cell
let validation = cell.DataValidation;

//Get the settings
let allowType = validation.AllowType.toString();
let data = validation.CompareOperator.toString();
let minimum = validation.Formula1.toString();
let maximum = validation.Formula2.toString();
let ignoreBlank = validation.IgnoreBlank.toString();

//Create an array to save the content
let content = [];

//Set string format for displaying
let result = `Settings of Validation: \r\nAllow Type: ${allowType}\r\nData: ${data}\r\nMinimum: ${minimum}\r\nMaximum: ${maximum}\r\nIgnoreBlank: ${ignoreBlank}`;

//Add result string to the content array
content.push(result);
```

---

# Excel List Data Validation
## Implement list data validation in Excel cells
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Set text for cells 
sheet.Range.get("A7").Text = "Beijing";
sheet.Range.get("A8").Text = "New York";
sheet.Range.get("A9").Text = "Denver";
sheet.Range.get("A10").Text = "Paris";

//Set data validation for cell
let range = sheet.Range.get("D10");
range.DataValidation.ShowError = true;
range.DataValidation.AlertStyle = wasmModule.AlertStyleType.Stop;
range.DataValidation.ErrorTitle = "Error";
range.DataValidation.ErrorMessage = "Please select a city from the list";
range.DataValidation.DataRange = sheet.Range.get("A7:A10");
```

---

# Remove Auto Filters in Excel
## This code demonstrates how to remove auto filters from an Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

//Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

//Remove the auto filters.
sheet.AutoFilters.Clear();

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# spire.xls javascript data validation
## remove data validation from Excel worksheet
```javascript
// Create a new workbook
const workbook = wasmModule.Workbook.Create();
// Load an existing Excel document
workbook.LoadFromFile({fileName: inputFileName});

//Create an array of rectangles, which is used to locate the ranges in worksheet.
let rectangles = [];

//Assign value to the first element of the array. This rectangle specifies the cells from A1 to B3.
rectangles.push(wasmModule.Rectangle.FromLTRB(0, 0, 1, 2));

//Remove validations in the ranges represented by rectangles.
workbook.Worksheets.get(0).DVTable.Remove(rectangles);
```

---

# Excel Data Validation Across Sheets
## Set data validation on cells that reference data from a different worksheet
```javascript
// Get the first worksheet
let sheet1 = workbook.Worksheets.get(0);

sheet1.Range.get("B10").Text = "Here is a dataValidation example.";

// This is the second sheet
let sheet2 = workbook.Worksheets.get(1);

// The property is to enable the data can be from different sheet
workbook.Allow3DRangesInDataValidation = true;
sheet1.Range.get("B11").DataValidation.DataRange = sheet2.Range.get("A1:A7");
```

---

# Excel Time Data Validation
## Set time validation rules for Excel cells
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

sheet.Range.get("C12").Text = "Please enter time between 09:00 and 18:00:";
sheet.Range.get("C12").AutoFitColumns();

//Set Time data validation for cell "D12"
let range = sheet.Range.get("D12");
range.DataValidation.AllowType = wasmModule.CellDataType.Time;
range.DataValidation.CompareOperator = wasmModule.ValidationComparisonOperator.Between;

range.DataValidation.Formula1 = "09:00";
range.DataValidation.Formula2 = "18:00";

range.DataValidation.AlertStyle = wasmModule.AlertStyleType.Info;
range.DataValidation.ShowError = true;
range.DataValidation.ErrorTitle = "Time Error";
range.DataValidation.ErrorMessage = "Please enter a valid time";
range.DataValidation.InputMessage = "Time Validation Type";
range.DataValidation.IgnoreBlank = true;
range.DataValidation.ShowInput = true;
```

---

# Excel Data Validation Verification
## Verify data against cell validation rules in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Cell B4 has the Decimal Validation
let cell = sheet.Range.get("B4");

// Get the validation of this cell
let validation = cell.DataValidation;

// Get the specified data range
let minimum = parseFloat(validation.Formula1);
let maximum = parseFloat(validation.Formula2);

// Create array to save results
let content = [];

// Set different numbers for the cell
for(let i=5; i<100; i+=40) {
    cell.NumberValue = i;
    let result = null;
    // Verify
    if(cell.NumberValue < minimum || cell.NumberValue > maximum) {
        // Set string format for displaying
        result = `Is input ${i} a valid value for this Cell: false`;
    } else {
        // Set string format for displaying
        result = `Is input ${i} a valid value for this Cell: true`;
    }
    // Add result string to array
    content.push(result);
}
```

---

# Excel Whole Number Data Validation
## Set whole number validation rules for Excel cells
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);
sheet.Range.get("C12").Text = "Please enter number between 10 and 100:";
sheet.Range.get("C12").AutoFitColumns();

//Set Whole Number data validation for cell "D12"
let range = sheet.Range.get("D12");
range.DataValidation.AllowType = wasmModule.CellDataType.Integer;
range.DataValidation.CompareOperator = wasmModule.ValidationComparisonOperator.Between;

range.DataValidation.Formula1 = "10";
range.DataValidation.Formula2 = "100";

range.DataValidation.AlertStyle = wasmModule.AlertStyleType.Info;
range.DataValidation.ShowError = true;
range.DataValidation.ErrorTitle = "Error";
range.DataValidation.ErrorMessage = "Please enter a valid number";
range.DataValidation.InputMessage = "Whole Number Validation Type";
range.DataValidation.IgnoreBlank = true;
range.DataValidation.ShowInput = true;
```

---

# Excel Chart Data Table
## Add a data table to a chart in Excel
```javascript
// Get the first worksheet from the workbook
let sheet = workbook.Worksheets.get(0);

// Get the first chart from the worksheet
let chart = sheet.Charts.get(0);

// Set the chart to display a data table
chart.HasDataTable = true;
```

---

# spire.xls javascript error bars
## add error bars to excel charts
```javascript
//Add a line chart and then add percentage error bar to the chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Line});
chart.DataRange = sheet.Range.get("B1:B7");
chart.SeriesDataFromRange = false;
//Set chart position
chart.TopRow = 8;
chart.BottomRow = 25;
chart.LeftColumn = 2;
chart.RightColumn = 9;
chart.ChartTitle = "Error Bar 10% Plus";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
let cs1 = chart.Series.get(0);
cs1.CategoryLabels = sheet.Range.get("A2:A7");
cs1.ErrorBar({bIsY:true, include:wasmModule.ErrorBarIncludeType.Plus, type:wasmModule.ErrorBarType.Percentage, numberValue:10.0});

// Add a column chart with standard error bars as comparison
let chart2 = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ColumnClustered});
chart2.DataRange = sheet.Range.get("B1:C7");
chart2.SeriesDataFromRange = false;

//Set chart position
chart2.TopRow = 8;
chart2.BottomRow = 25;
chart2.LeftColumn = 10;
chart2.RightColumn = 17;
chart2.ChartTitle = "Standard Error Bar";
chart2.ChartTitleArea.IsBold = true;
chart2.ChartTitleArea.Size = 12;
let cs2 = chart2.Series.get(0);
cs2.CategoryLabels = sheet.Range.get("A2:A7");
cs2.ErrorBar({bIsY:true, include:wasmModule.ErrorBarIncludeType.Minus, type:wasmModule.ErrorBarType.StandardError, numberValue:0.3});
let cs3 = chart2.Series.get(1);
cs3.ErrorBar({bIsY:true, include:wasmModule.ErrorBarIncludeType.Both, type:wasmModule.ErrorBarType.StandardError, numberValue:0.5});
```

---

# spire.xls javascript chart
## add picture to chart
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Get the chart
let chart = sheet.Charts.get(0);

//Add the picture in chart
chart.Shapes.AddPicture("SpireXls.png");
```

---

# spire.xls javascript textbox
## add textbox to chart in excel worksheet
```javascript
//Get the first chart
let chart = sheet.Charts.get(0);

//Add a Textbox
let textbox = chart.Shapes.AddTextBox();
textbox.Width = 1200;
textbox.Height = 320;
textbox.Left = 1000;
textbox.Top = 480;
textbox.Text = "This is a textbox";
```

---

# spire.xls javascript chart trendlines
## add different types of trendlines to excel charts
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//select chart and set logarithmic trendline
let chart = sheet.Charts.get(0);
chart.ChartTitle = "Logarithmic Trendline";
chart.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Logarithmic});

//select chart and set moving_average trendline
let chart1 = sheet.Charts.get(1);
chart1.ChartTitle = "Moving Average Trendline";
chart1.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Moving_Average});

//select chart and set linear trendline
let chart2 = sheet.Charts.get(2);
chart2.ChartTitle = "Linear Trendline";
chart2.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Linear});

//select chart and set exponential trendline
let chart3 = sheet.Charts.get(3);
chart3.ChartTitle = "Exponential Trendline";
chart3.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Exponential});
```

---

# spire.xls javascript chart
## adjust bar space in chart
```javascript
//Get the first worksheet from workbook and then get the first chart from the worksheet
let ws = workbook.Worksheets.get(0);
let chart = ws.Charts.get(0);

//Adjust the space between bars
for(let cs of chart.Series) {
    cs.Format.Options.GapWidth = 200;
    cs.Format.Options.Overlap = 0;
}
```

---

# Spire.XLS JavaScript Chart Soft Edges Effect
## Apply soft edges effect to Excel chart
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Get the chart
let chart = sheet.Charts.get(0);

//Specify the size of the soft edge. Value can be set from 0 to 100
chart.ChartArea.Shadow.SoftEdge = 25;
```

---

# spire.xls javascript chart
## change chart size and position
```javascript
//Get the chart
let chart = sheet.Charts.get(0);

//Change chart size
chart.Width = 600;
chart.Height = 500;

//Change chart position
chart.LeftColumn = 3;
chart.TopRow = 7;
```

---

# Chart Data Label Modification
## Change data label text in Excel chart
```javascript
// Get the chart
let chart = sheet.Charts.get(0);

// Change data label of the first datapoint of the first series
chart.Series.get(0).DataPoints.get(0).DataLabels.Text = "changed data label";
```

---

# Spire.XLS JavaScript Chart
## Change Data Range of Chart
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Get chart
let chart = sheet.Charts.get(0);

//Change data range
chart.DataRange = sheet.Range.get("A1:C4");
```

---

# spire.xls javascript chart
## change major gridlines color in chart
```javascript
// Get the chart from the first worksheet
let sheet = workbook.Worksheets.get(0);
let chart = sheet.Charts.get(0);

// Change the color of major gridlines
chart.PrimaryValueAxis.MajorGridLines.LineProperties.Color = wasmModule.Color.get_Red();
```

---

# Excel Chart Series Color Change
## Change the color of a chart series in Excel
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Get the first chart
let chart = sheet.Charts.get(0);

//Get the second series
let cs = chart.Series.get(1);

//Set the fill type
cs.Format.Fill.FillType = wasmModule.ShapeFillType.SolidColor;

//Change the fill color
cs.Format.Fill.ForeColor = wasmModule.Color.get_Orange();
```

---

# Excel Chart Axis Title
## Set titles for chart axes in Excel
```javascript
//Get the chart
let chart = sheet.Charts.get(0);

//Set axis title
chart.PrimaryCategoryAxis.Title = "Category Axis";
chart.PrimaryValueAxis.Title = "Value axis";

//Set font size
chart.PrimaryCategoryAxis.Font.Size = 12;
chart.PrimaryValueAxis.Font.Size = 12;
```

---

# spire.xls javascript chart to image
## convert excel chart to image
```javascript
// Create a new workbook object
const workbook = wasmModule.Workbook.Create();
// Load the Excel file 
workbook.LoadFromFile(excelFileName);

//Save chart as image
const image = workbook.SaveChartAsImage({worksheet:workbook.Worksheets.get(0), chartIndex:0});
const outputFile = 'ChartToImage.png';
image.Save(outputFile);

// Dispose of the workbook object to free resources
workbook.Dispose();
```

---

# spire.xls javascript chart
## create clustered bar chart
```javascript
// Add a chart
let chart = sheet.Charts.Add();

// Set region of chart data
chart.DataRange = sheet.Range.get("A1:C5");
chart.SeriesDataFromRange = false;

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;
chart.ChartType = wasmModule.ExcelChartType.BarClustered;

// Chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

chart.PrimaryCategoryAxis.Title = "Country";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.TextRotationAngle = 90;

chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;

for (let cs of chart.Series) {
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# spire.xls javascript chart
## create 3D clustered bar chart
```javascript
// Add a chart
let chart = sheet.Charts.Add();

// Set region of chart data
chart.DataRange = sheet.Range.get("A1:C5");
chart.SeriesDataFromRange = false;

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;
chart.ChartType = wasmModule.ExcelChartType.Bar3DClustered;

// Chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

chart.PrimaryCategoryAxis.Title = "Country";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.TextRotationAngle = 90;

chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;

for (let cs of chart.Series) {
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# spire.xls javascript chart
## create clustered column chart
```javascript
// Add a new chart to the worksheet
const chart = sheet.Charts.Add();
chart.DataRange = sheet.Range.get("A1:C5");
chart.SeriesDataFromRange = false;

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

// Set chart type to clustered column
chart.ChartType = wasmModule.ExcelChartType.ColumnClustered;

// Configure chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Configure category axis
chart.PrimaryCategoryAxis.Title = "Country";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

// Configure value axis
chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;

// Configure series
for (let i = 0; i < chart.Series.Length; i++) {
    let cs = chart.Series.get(i);
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

// Set legend position
chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# spire.xls javascript chart
## create Box and Whisker chart
```javascript
// Add a new chart
let officeChart = sheet.Charts.Add();

// Set the chart title
officeChart.ChartTitle = "Yearly Vehicle Sales";

// Set chart type as Box and Whisker
officeChart.ChartType = wasmModule.ExcelChartType.BoxAndWhisker;

// Set data range in the worksheet
officeChart.DataRange = sheet.Range.get("A1:E17");

// Box and Whisker settings on first series
let seriesA = officeChart.Series.get(0);
seriesA.DataFormat.ShowInnerPoints = false;
seriesA.DataFormat.ShowOutlierPoints = true;
seriesA.DataFormat.ShowMeanMarkers = true;
seriesA.DataFormat.ShowMeanLine = false;
seriesA.DataFormat.QuartileCalculationType = wasmModule.ExcelQuartileCalculation.ExclusiveMedian;

// Box and Whisker settings on second series
let seriesB = officeChart.Series.get(1);
seriesB.DataFormat.ShowInnerPoints = false;
seriesB.DataFormat.ShowOutlierPoints = true;
seriesB.DataFormat.ShowMeanMarkers = true;
seriesB.DataFormat.ShowMeanLine = false;
seriesB.DataFormat.QuartileCalculationType = wasmModule.ExcelQuartileCalculation.InclusiveMedian;

// Box and Whisker settings on third series
let seriesC = officeChart.Series.get(2);
seriesC.DataFormat.ShowInnerPoints = false;
seriesC.DataFormat.ShowOutlierPoints = true;
seriesC.DataFormat.ShowMeanMarkers = true;
seriesC.DataFormat.ShowMeanLine = false;
seriesC.DataFormat.QuartileCalculationType = wasmModule.ExcelQuartileCalculation.ExclusiveMedian;
```

---

# spire.xls javascript chart
## create Bubble chart
```javascript
// Add a chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Bubble});

// Set region of chart data
chart.DataRange = sheet.Range.get("A1:C5");
chart.SeriesDataFromRange = false;
chart.Series.get(0).Bubbles = sheet.Range.get("C2:C5");

// Set position of chart
chart.LeftColumn = 7;
chart.TopRow = 6;
chart.RightColumn = 16;
chart.BottomRow = 29;

chart.ChartTitle = "Bubble Chart";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
```

---

# Spire.XLS JavaScript Chart
## Create chart based on pivot table
```javascript
// Create a new workbook object
const workbook = wasmModule.Workbook.Create();

// Load the Excel file 
workbook.LoadFromFile(excelFileName);

// Get the sheet in which the pivot table is located
let sheet = workbook.Worksheets.get(0);

let pt = sheet.PivotTables.get(0);

workbook.Worksheets.get(1).Charts.Add({pivotChartType: wasmModule.ExcelChartType.BarClustered, pivotTable: pt});

// Dispose of the workbook object to free resources
workbook.Dispose();
```

---

# Excel Chart Creation Without Range Data
## Create a chart in Excel using direct values instead of range data
```javascript
// Create a new workbook object
const workbook = wasmModule.Workbook.Create();
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Add a chart to the worksheet
let chart = sheet.Charts.Add();
chart.ChartTitle = "Sample Chart";

// Add a series to the chart
let series = chart.Series.Add();

// Add data directly without using range data
series.EnteredDirectlyValues = [wasmModule.Int32.Create(10), wasmModule.Int32.Create(20), wasmModule.Int32.Create(30)];
```

---

# spire.xls javascript custom chart
## create custom chart with different chart types for different series
```javascript
// Add a chart based on the data from A1 to B4
let chart = sheet.Charts.Add();
chart.DataRange = sheet.Range.get("A1:B4");
chart.SeriesDataFromRange = false;

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 10;
chart.RightColumn = 7;
chart.BottomRow = 25;

// Apply different chart type to different series
let cs1 = chart.Series.get(0);
cs1.SerieType = wasmModule.ExcelChartType.ColumnClustered;
let cs2 = chart.Series.get(1);
cs2.SerieType = wasmModule.ExcelChartType.Line;

chart.ChartTitle = "Custom chart";
```

---

# spire.xls javascript chart
## create doughnut chart
```javascript
// Add a new chart, set chart type as doughnut
let chart = sheet.Charts.Add();
chart.ChartType = wasmModule.ExcelChartType.Doughnut;
chart.DataRange = sheet.Range.get("A1:B5");
chart.SeriesDataFromRange = false;

// Set position of chart
chart.LeftColumn = 4;
chart.TopRow = 2;
chart.RightColumn = 12;
chart.BottomRow = 22;

// Chart title
chart.ChartTitle = "Market share by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

for (let i = 0; i < chart.Series.Count; i++) {
    chart.Series.get(i).DataPoints.DefaultDataPoint.DataLabels.HasPercentage = true;
}

chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# spire.xls javascript funnel chart
## create funnel chart in Excel
```javascript
//Add a new chart
let officeChart = sheet.Charts.Add();
//Set chart type as Funnel
officeChart.ChartType = wasmModule.ExcelChartType.Funnel;

//Set data range in the worksheet
officeChart.DataRange = sheet.Range.get("A1:B6");

//Set the chart title
officeChart.ChartTitle = "Funnel";

//Formatting the legend and data label option
officeChart.HasLegend = false;
officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;
```

---

# spire.xls javascript histogram chart
## create histogram chart in excel
```javascript
//Add a new chart
let officeChart = sheet.Charts.Add();
//Set chart type as histogram
officeChart.ChartType = wasmModule.ExcelChartType.Histogram;

//Set data range in the worksheet
officeChart.DataRange = sheet.Range.get("A1:A15");
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 12;

//Category axis bin settings
officeChart.PrimaryCategoryAxis.BinWidth = 8;

//Gap width settings
officeChart.Series.get(0).DataFormat.Options.GapWidth = 6;

//Set the chart title and axis title
officeChart.ChartTitle = "Height Data";
officeChart.PrimaryValueAxis.Title = "Number of students";
officeChart.PrimaryCategoryAxis.Title = "Height";

//Hiding the legend
officeChart.HasLegend = false;
```

---

# spire.xls javascript chart
## create multi-level category chart
```javascript
//Add a clustered bar chart to worksheet
let chart = sheet.Charts.Add({chartType: wasmModule.ExcelChartType.BarClustered});
chart.ChartTitle = "Value";
chart.PlotArea.Fill.FillType = wasmModule.ShapeFillType.NoFill;
chart.Legend.Delete();
chart.LeftColumn = 5;
chart.TopRow = 1;
chart.RightColumn = 14;

//Set the data source of series data
chart.DataRange = sheet.Range.get("C2:C9");
chart.SeriesDataFromRange = false;
//Set the data source of category labels
let serie = chart.Series.get(0);
serie.CategoryLabels = sheet.Range.get("A2:B9");
//Show multi-level category labels
chart.PrimaryCategoryAxis.MultiLevelLable = true;
```

---

# spire.xls javascript chart
## create Pareto chart
```javascript
//Add chart
let officeChart = sheet.Charts.Add();
//Set chart type as Pareto
officeChart.ChartType = wasmModule.ExcelChartType.Pareto;

//Set data range in the worksheet
officeChart.DataRange = sheet.Range.get("A2:B8");

officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 12;
officeChart.PrimaryCategoryAxis.IsBinningByCategory = true;

officeChart.PrimaryCategoryAxis.OverflowBinValue = 5;
officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1;

//Formatting Pareto line
officeChart.Series.get(0).ParetoLineFormat.LineProperties.Color = wasmModule.Color.get_Blue();

//Gap width settings
officeChart.Series.get(0).DataFormat.Options.GapWidth = 6;

//Set the chart title
officeChart.ChartTitle = "Expenses";
```

---

# spire.xls javascript pivot chart
## create a pivot chart based on an existing pivot table
```javascript
//get the first worksheet
let sheet = workbook.Worksheets.get(0);
//get the first pivot table in the worksheet
let pivotTable = sheet.PivotTables.get(0);

//create a clustered column chart based on the pivot table
let chart = sheet.Charts.Add({pivotChartType: wasmModule.ExcelChartType.ColumnClustered, pivotTable: pivotTable});
//set chart position
chart.TopRow = 10;
chart.LeftColumn = 1;
chart.RightColumn = 7;
chart.BottomRow = 25;
//set chart title
chart.ChartTitle = "Pivot Chart";
```

---

# Spire.XLS JavaScript Radar Chart
## Create a radar chart in an Excel workbook

```javascript
// Add a new chart worksheet to workbook
let chart = sheet.Charts.Add();

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

// Set region of chart data
chart.DataRange = sheet.Range.get("A1:C5");
chart.SeriesDataFromRange = false;

chart.ChartType = wasmModule.ExcelChartType.Radar;

// Chart title
chart.ChartTitle = "Sale market by region";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
chart.PlotArea.Fill.Visible = false;
chart.Legend.Position = wasmModule.LegendPositionType.Corner;
```

---

# spire.xls javascript chart
## create SunBurst Chart
```javascript
//Add chart
let officeChart = sheet.Charts.Add();
//Set chart type as Sunburst
officeChart.ChartType = wasmModule.ExcelChartType.SunBurst;

//Set data range in the worksheet
officeChart.DataRange = sheet.Range.get("A1:D16");

officeChart.TopRow = 1;
officeChart.BottomRow = 17;
officeChart.LeftColumn = 6;
officeChart.RightColumn = 14;

//Set the chart title
officeChart.ChartTitle = "Sales by quarter";

//Formatting data labels
officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;

//Hiding the legend
officeChart.HasLegend = false;
```

---

# spire.xls javascript treemap chart
## create TreeMap chart in Excel
```javascript
//Add chart
let officeChart = sheet.Charts.Add();
//Set chart type as TreeMap
officeChart.ChartType = wasmModule.ExcelChartType.TreeMap;

//Set data range in the worksheet
officeChart.DataRange = sheet.Range.get("A2:C11");
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 14;

//Set the chart title
officeChart.ChartTitle = "Area by countries";

//Set the Treemap label option
officeChart.Series.get(0).DataFormat.TreeMapLabelOption = wasmModule.ExcelTreeMapLabelOption.Banner;

//Formatting data labels      
officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;
```

---

# spire.xls javascript waterfall chart
## create Waterfall chart in Excel
```javascript
// Add a new chart to worksheet
let officeChart = sheet.Charts.Add();
// Set chart type as waterfall
officeChart.ChartType = wasmModule.ExcelChartType.WaterFall;

// Set data range to the chart from the worksheet
officeChart.DataRange = sheet.Range.get("A2:B8");

// Set chart position and size
officeChart.TopRow = 1;
officeChart.BottomRow = 19;
officeChart.LeftColumn = 4;
officeChart.RightColumn = 12;

// Data point settings as total in chart
officeChart.Series.get(0).DataPoints.get(3).SetAsTotal = true;
officeChart.Series.get(0).DataPoints.get(6).SetAsTotal = true;

// Showing the connector lines between data points
officeChart.Series.get(0).Format.ShowConnectorLines = true;

// Set the chart title
officeChart.ChartTitle = "WaterFall Chart";

// Formatting data label and legend option
officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;
officeChart.Legend.Position = wasmModule.LegendPositionType.Right;
```

---

# Custom Data Markers for Excel Charts
## Create and customize data markers in scatter charts
```javascript
//Create a Scatter-Markers chart based on the sample data
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ScatterMarkers});
chart.DataRange = sheet.Range.get("A1:B7");
chart.PlotArea.Visible = false;
chart.SeriesDataFromRange = false;
chart.TopRow = 5;
chart.BottomRow = 22;
chart.LeftColumn = 4;
chart.RightColumn = 11;
chart.ChartTitle = "Chart with Markers";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 10;

//Format the markers in the chart by setting the background color, foreground color, type, size and transparency
let cs1 = chart.Series.get(0);
cs1.DataFormat.MarkerBackgroundColor = wasmModule.Color.get_RoyalBlue();
cs1.DataFormat.MarkerForegroundColor = wasmModule.Color.get_WhiteSmoke();
cs1.DataFormat.MarkerSize = 7;
cs1.DataFormat.MarkerStyle = wasmModule.ChartMarkerType.PlusSign;
cs1.DataFormat.MarkerTransparencyValue = 0.8;

let cs2 = chart.Series.get(1);
cs2.DataFormat.MarkerBackgroundColor = wasmModule.Color.get_Pink();
cs2.DataFormat.MarkerSize = 9;
cs2.DataFormat.MarkerStyle = wasmModule.ChartMarkerType.Triangle;
cs2.DataFormat.MarkerTransparencyValue = 0.9;
```

---

# Excel Chart Data Callout Configuration
## Set data callout properties for chart data labels
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);

//Get the first chart
let chart = sheet.Charts.get(0);

for(let cs of chart.Series) {
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasWedgeCallout = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = true;
}
```

---

# Excel Chart Legend Management
## Delete specific legend entries from an Excel chart
```javascript
//Get the chart
let chart = sheet.Charts.get(0);

//Delete the first and the second legend entries from the chart
chart.Legend.LegendEntries.get(0).Delete();
chart.Legend.LegendEntries.get(1).Delete();
```

---

# Excel Chart with Discontinuous Data
## Create a chart with discontinuous data ranges for series
```javascript
//Add a chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ColumnClustered});
chart.SeriesDataFromRange = false;

//Set the position of chart
chart.LeftColumn = 1;
chart.TopRow = 10;
chart.RightColumn = 10;
chart.BottomRow = 24;

//Add a series
let cs1 = chart.Series.Add();

//Set the name of the cs1
cs1.Name = sheet.Range.get("B1").Value;

//Set discontinuous values for cs1
cs1.CategoryLabels = sheet.Range.get("A2:A3").AddCombinedRange(sheet.Range.get("A5:A6")).AddCombinedRange(sheet.Range.get("A8:A9"));
cs1.Values = sheet.Range.get("B2:B3").AddCombinedRange(sheet.Range.get("B5:B6")).AddCombinedRange(sheet.Range.get("B8:B9"));

//Set the chart type
cs1.SerieType = wasmModule.ExcelChartType.ColumnClustered;

//Add a series
let cs2 = chart.Series.Add();
cs2.Name = sheet.Range.get("C1").Value;
cs2.CategoryLabels = sheet.Range.get("A2:A3").AddCombinedRange(sheet.Range.get("A5:A6")).AddCombinedRange(sheet.Range.get("A8:A9"));
cs2.Values = sheet.Range.get("C2:C3").AddCombinedRange(sheet.Range.get("C5:C6")).AddCombinedRange(sheet.Range.get("C8:C9"));
cs2.SerieType = wasmModule.ExcelChartType.ColumnClustered;

chart.ChartTitle = "Chart";
chart.ChartTitleArea.Font.Size = 20;
chart.ChartTitleArea.Color = wasmModule.Color.get_Black();

chart.PrimaryValueAxis.HasMajorGridLines = false;
```

---

# spire.xls javascript chart
## edit line chart by adding new series
```javascript
// Get the line chart
let chart = sheet.Charts.get(0);

// Add a new series
let cs = chart.Series.Add({name:"Added"});

// Set the values for the series
cs.Values = sheet.Range.get("I1:L1");
```

---

# Exploded Doughnut Chart Creation
## Create an exploded doughnut chart in Excel using JavaScript
```javascript
//Add a chart
let chart = sheet.Charts.Add();
chart.ChartType = wasmModule.ExcelChartType.DoughnutExploded;

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

//Set region of chart data
chart.DataRange = sheet.Range.get("A1:B5");
chart.SeriesDataFromRange = false;

//Chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

for(let cs of chart.Series) {
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

chart.PlotArea.Fill.Visible = false;
chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# Extract Trendline Equation from Chart
## This code demonstrates how to extract the trendline equation from a chart in an Excel file

```javascript
//Get the chart from the first worksheet
let chart = workbook.Worksheets.get(0).Charts.get(0);

//Get the trendline of the chart and then extract the equation of the trendline
let trendLine = chart.Series.get(1).TrendLines.get(0);
let formula = trendLine.Formula;
let equation = "The equation is: " + formula;
```

---

# Fill Chart Elements with Picture
## Fill chart plot area with a custom image in Excel using JavaScript
```javascript
let ws = workbook.Worksheets.get(0);
let chart = ws.Charts.get(0);

// Fill plot area with image
chart.PlotArea.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile("Background.png"), name:"None"});
```

---

# spire.xls javascript chart axis formatting
## format chart axis properties and appearance
```javascript
//Add a chart
let chart = sheet.Charts.Add({chartType: wasmModule.ExcelChartType.ColumnClustered});
chart.DataRange = sheet.Range.get("B1:B9");
chart.SeriesDataFromRange = false;
chart.PlotArea.Visible = false;
chart.TopRow = 10;
chart.BottomRow = 28;
chart.LeftColumn = 2;
chart.RightColumn = 10;
chart.ChartTitle = "Chart with Customized Axis";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
let cs1 = chart.Series.get(0);
cs1.CategoryLabels = sheet.Range.get("A2:A9");

//Format axis
chart.PrimaryValueAxis.MajorUnit = 8;
chart.PrimaryValueAxis.MinorUnit = 2;
chart.PrimaryValueAxis.MaxValue = 50;
chart.PrimaryValueAxis.MinValue = 0;
chart.PrimaryValueAxis.IsReverseOrder = false;
chart.PrimaryValueAxis.MajorTickMark = wasmModule.TickMarkType.TickMarkOutside;
chart.PrimaryValueAxis.MinorTickMark = wasmModule.TickMarkType.TickMarkInside;
chart.PrimaryValueAxis.TickLabelPosition = wasmModule.TickLabelPositionType.TickLabelPositionNextToAxis;
chart.PrimaryValueAxis.CrossesAt = 0;

//Set NumberFormat
chart.PrimaryValueAxis.NumberFormat = "$#,##0";
chart.PrimaryValueAxis.IsSourceLinked = false;

let serie = chart.Series.get(0);
for(let dataPoint of serie.DataPoints) {
    //Format Series
    dataPoint.DataFormat.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
    dataPoint.DataFormat.Fill.ForeColor = Module.wasmModule.Color.get_LightGreen();

    //Set transparency
    dataPoint.DataFormat.Fill.Transparency = 0.3;
}
```

---

# spire.xls javascript gauge chart
## create gauge chart using doughnut and pie chart combination
```javascript
//Add a Doughnut chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Doughnut});
chart.DataRange = sheet.Range.get("A1:A5");
chart.SeriesDataFromRange = false;
chart.HasLegend = true;

//Set the position of chart
chart.LeftColumn = 2;
chart.TopRow = 7;
chart.RightColumn = 9;
chart.BottomRow = 25;

//Get the series 1
let cs1 = chart.Series.get({name:"Value"});
cs1.Format.Options.DoughnutHoleSize = 60;
cs1.DataFormat.Options.FirstSliceAngle = 270;

//Set the fill color
cs1.DataPoints.get(0).DataFormat.Fill.ForeColor = wasmModule.Color.get_Yellow();
cs1.DataPoints.get(1).DataFormat.Fill.ForeColor = wasmModule.Color.get_PaleVioletRed
cs1.DataPoints.get(2).DataFormat.Fill.ForeColor = wasmModule.Color.get_DarkViolet();
cs1.DataPoints.get(3).DataFormat.Fill.Visible = false;

//Add a series with pie chart
let cs2 = chart.Series.Add({name:"Pointer", serieType:wasmModule.ExcelChartType.Pie});

//Set the value
cs2.Values = sheet.Range.get("D2:D4");
cs2.UsePrimaryAxis = false;
cs2.DataPoints.get(0).DataLabels.HasValue = true;
cs2.DataFormat.Options.FirstSliceAngle = 270;
cs2.DataPoints.get(0).DataFormat.Fill.Visible = false;
cs2.DataPoints.get(1).DataFormat.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
cs2.DataPoints.get(1).DataFormat.Fill.ForeColor = wasmModule.Color.get_Black();
cs2.DataPoints.get(2).DataFormat.Fill.Visible = false;
```

---

# spire.xls javascript chart
## get category labels from chart
```javascript
//Create a workbook
workbook.LoadFromFile(excelFileName);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Get the chart
let chart = sheet.Charts.get(0);

//Get the cell range of the category labels
let cr = chart.PrimaryCategoryAxis.CategoryLabels;
for (let i = 0; i < cr.Count; i++) {
    sb.push(cr.Cells.get(i).Value);
}
```

---

# Spire.XLS JavaScript Chart Data Point Values
## Extract values from chart data points in Excel
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);

// Get the chart
let chart = sheet.Charts.get(0);

// Get the first series of the chart
let cs = chart.Series.get(0);

for(let cr of cs.Values.Cells) {
    // Get the range address
    sb.push(cr.RangeAddress);

    // Get the data point value
    sb.push(`The value of the data point is ${cr.Value}`);
}
```

---

# spire.xls javascript chart
## get worksheet of a chart
```javascript
// Access first worksheet of the workbook
let worksheet = workbook.Worksheets.get(0);

// Access the first chart inside this worksheet
let chart = worksheet.Charts.get(0);

// Get its worksheet
let obj = chart.Worksheet;
let wSheet = wasmModule.Worksheet.Convert(obj);

// Set string format for displaying
let result = `Sheet Name: ${worksheet.Name}\r\nCharts' sheet Name: ${wSheet.Name}`;
```

---

# Hide Major Gridlines in Chart
## This code demonstrates how to hide major gridlines in a chart using Spire.XLS for JavaScript
```javascript
let sheet = workbook.Worksheets.get(0);

//Get the chart
let chart = sheet.Charts.get(0);

//Hide major gridlines
chart.PrimaryValueAxis.HasMajorGridLines = false;
```

---

# spire.xls javascript chart
## create Line chart
```javascript
//Add a chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Line});

//Set region of chart data
chart.DataRange = sheet.Range.get("A1:E5");

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

//Set chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

chart.PrimaryCategoryAxis.Title = "Month";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;

for(let cs of chart.Series) {
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

chart.PlotArea.Fill.Visible = false;

chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# spire.xls javascript chart
## create Pie chart
```javascript
let sheet = workbook.Worksheets.get(0);
sheet.Name = "Pie Chart";

//Add a chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Pie});

//Set region of chart data
chart.DataRange = sheet.Range.get("B2:B5");
chart.SeriesDataFromRange = false;

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 9;
chart.BottomRow = 25;

//Chart title
chart.ChartTitle = "Sales by year";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

let cs = chart.Series.get(0);
cs.CategoryLabels = sheet.Range.get("A2:A5");
cs.Values = sheet.Range.get("B2:B5");
cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;

chart.PlotArea.Fill.Visible = false;
```

---

# spire.xls javascript chart
## create pyramid column chart
```javascript
//Add a chart
let chart = sheet.Charts.Add();

//Set region of chart data
chart.DataRange = sheet.Range.get("B2:B5");
chart.SeriesDataFromRange = false;

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;

chart.ChartType = wasmModule.ExcelChartType.Pyramid3DClustered;

//Chart title
chart.ChartTitle = "Sales by year";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

chart.PrimaryCategoryAxis.Title = "Year";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;

let cs = chart.Series.get(0);
cs.CategoryLabels = sheet.Range.get("A2:A5");
cs.Format.Options.IsVaryColor = true;
```

---

# spire.xls javascript chart
## remove chart from excel worksheet
```javascript
//Get the first worksheet from the workbook
let sheet = workbook.Worksheets.get(0);
//Get the first chart from the first worksheet
let chart = sheet.Charts.get(0);
//Remove the chart
chart.Remove();
```

---

# Excel Chart Resizing and Moving
## Resize and reposition an Excel chart in a worksheet
```javascript
// Get the chart from the first worksheet
let sheet = workbook.Worksheets.get(0);
let chart = sheet.Charts.get(0);

// Set position of the chart
chart.LeftColumn = 5;
chart.TopRow = 1;

// Resize the chart
chart.Width = 500;
chart.Height = 350;
```

---

# Excel Chart Data Labels Rich Text
## Setting rich text formatting for data labels in Excel charts
```javascript
//Get the first datalabel of the first series
let datalabel = chart.Series.get(0).DataPoints.get(0).DataLabels;

//Set the text
datalabel.Text = "Rich Text Label";

//Show the value
chart.Series.get(0).DataPoints.get(0).DataLabels.HasValue = true;

//Set styles for the text
chart.Series.get(0).DataPoints.get(0).DataLabels.Color = wasmModule.Color.get_Red();
chart.Series.get(0).DataPoints.get(0).DataLabels.IsBold = true;
```

---

# Excel 3D Chart Rotation
## Set X and Y rotation angles for a 3D chart in Excel
```javascript
//Get the chart from the first worksheet
let sheet = workbook.Worksheets.get(0);
let chart = sheet.Charts.get(0);

//X rotation:
chart.Rotation = 30;
//Y rotation:
chart.Elevation = 20;
```

---

# spire.xls javascript chart
## create Scatter Chart
```javascript
//Add a chart
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ScatterMarkers});

//Set region of chart data
chart.DataRange = sheet.Range.get("B2:B10");
chart.SeriesDataFromRange = false;

//Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 11;
chart.RightColumn = 10;
chart.BottomRow = 28;

chart.ChartTitle = "Scatter Chart";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

chart.Series.get(0).CategoryLabels = sheet.Range.get("A2:A10");
chart.Series.get(0).Values = sheet.Range.get("B2:B10");

//Add a trend line for the first series
chart.Series.get(0).TrendLines.Add({type:wasmModule.TrendLineType.Exponential});

chart.PrimaryValueAxis.Title = "Salary";
chart.PrimaryCategoryAxis.Title = "Car Price";
```

---

# spire.xls javascript chart
## set and format data labels for charts
```javascript
let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.LineMarkers});
chart.DataRange = sheet.Range.get("B1:B7");
chart.PlotArea.Visible = false;
chart.SeriesDataFromRange = false;
chart.TopRow = 5;
chart.BottomRow = 26;
chart.LeftColumn = 2;
chart.RightColumn = 11;
chart.ChartTitle = "Data Labels Demo";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;
let cs1 = chart.Series.get(0);
cs1.CategoryLabels = sheet.Range.get("A2:A7");

cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = false;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = false;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". ";

cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9;
cs1.DataPoints.DefaultDataPoint.DataLabels.Color = wasmModule.Color.get_Red();
cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri";
cs1.DataPoints.DefaultDataPoint.DataLabels.Position = wasmModule.DataLabelPositionType.Center;
```

---

# spire.xls javascript chart
## set border color and style for chart
```javascript
//Get the first worksheet from workbook and then get the first chart from the worksheet
let ws = workbook.Worksheets.get(0);
let chart = ws.Charts.get(0);

//Set CustomLineWeight property for Series line
chart.Series.get(0).DataPoints.get(0).DataFormat.LineProperties.CustomLineWeight = 2.5;
//Set color property for Series line
chart.Series.get(0).DataPoints.get(0).DataFormat.LineProperties.Color = wasmModule.Color.get_Red();
```

---

# Setting Border Width of Chart Markers
## Adjust the border width of markers in Excel chart series
```javascript
// Get the chart from the first worksheet
let chart = workbook.Worksheets.get(0).Charts.get(0);

// Set marker border width for different series (unit is pt)
chart.Series.get(0).DataFormat.MarkerBorderWidth = 1.5;
chart.Series.get(1).DataFormat.MarkerBorderWidth = 2.5;
```

---

# Chart Background Color Setting
## Set background color for Excel chart
```javascript
// Get the first worksheet from workbook and then get the first chart from the worksheet
let ws = workbook.Worksheets.get(0);
let chart = ws.Charts.get(0);

// Set background color
chart.ChartArea.ForeGroundColor = wasmModule.Color.get_LightYellow();
```

---

# Set Color for Chart Area
## Set the foreground color for chart area and plot area in Excel
```javascript
// Get the chart
let chart = sheet.Charts.get(0);

// Set color for chart area
chart.ChartArea.Fill.ForeColor = wasmModule.Color.get_LightSeaGreen();

// Set color for plot area
chart.PlotArea.Fill.ForeColor = wasmModule.Color.get_LightGray();
```

---

# spire.xls javascript chart font
## set font for chart datapoints
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);

// Get the first chart
let chart = sheet.Charts.get(0);

// Create a font
let font = workbook.CreateFont();
font.Size = 15.0;
font.Color = wasmModule.Color.get_LightSeaGreen();

for (let cs of chart.Series) {
    // Set font
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
}
```

---

# Spire.XLS JavaScript Chart Font Settings
## Set font for chart legend and data table
```javascript
// Get the first worksheet from workbook
let ws = workbook.Worksheets.get(0);
let chart = ws.Charts.get(0);

// Create a font with specified size and color
let font = workbook.CreateFont();
font.Size = 14.0;
font.Color = wasmModule.Color.get_Red();

// Apply the font to chart Legend
chart.Legend.TextArea.SetFont(font);

// Apply the font to chart DataLabel
for (let cs of chart.Series) {
    cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
}
```

---

# Excel Chart Font Formatting
## Set font styles for chart title and axis
```javascript
// Set font for chart title and chart axis
let worksheet = workbook.Worksheets.get(0);
let chart = worksheet.Charts.get(0);

// Format the font for the chart title
chart.ChartTitleArea.Color = wasmModule.Color.get_Blue();
chart.ChartTitleArea.Size = 20.0;

// Format the font for the chart Axis
chart.PrimaryValueAxis.Font.Color = wasmModule.Color.get_Gold();
chart.PrimaryValueAxis.Font.Size = 10.0;
chart.PrimaryCategoryAxis.Font.Color = wasmModule.Color.get_Red();
chart.PrimaryCategoryAxis.Font.Size = 20.0;
```

---

# spire.xls javascript chart
## set legend background color
```javascript
let ws = workbook.Worksheets.get(0);
let chart = ws.Charts.get(0);

let x = chart.Legend.FrameFormat;
x.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
x.ForeGroundColor = wasmModule.Color.get_SkyBlue();
```

---

# Excel Chart Trendline Number Format
## Set number format for chart trendline in Excel
```javascript
// Get the chart from the first worksheet
let chart = workbook.Worksheets.get(0).Charts.get(0);

// Get the trendline of the chart and then extract the equation of the trendline
let trendLine = chart.Series.get(1).TrendLines.get(0);

// Set the number format of trendLine to "#,##0.00"
trendLine.DataLabel.NumberFormat = "#,##0.00";
```

---

# spire.xls javascript chart
## show leader lines for data labels in chart
```javascript
// Add a chart with BarStacked type
let chart = sheet.Charts.Add({ chartType: wasmModule.ExcelChartType.BarStacked });
chart.DataRange = sheet.Range.get("A1:C3");
chart.TopRow = 4;
chart.LeftColumn = 2;
chart.Width = 450;
chart.Height = 300;

// Show leader lines for data labels in chart series
for (let cs of chart.Series) {
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;
}
```

---

# spire.xls javascript sparkline
## create sparkline charts in excel
```javascript
// Add sparkline
let sparklineGroup = sheet.SparklineGroups.AddGroup({sparklineType:wasmModule.SparklineType.Line});
let sparklines = sparklineGroup.Add();
sparklines.Add({dataRange:sheet.Range.get("A2:D2"), referenceRange:sheet.Range.get("E2")});
sparklines.Add({dataRange:sheet.Range.get("A3:D3"), referenceRange:sheet.Range.get("E3")});
sparklines.Add({dataRange:sheet.Range.get("A4:D4"), referenceRange:sheet.Range.get("E4")});
sparklines.Add({dataRange:sheet.Range.get("A5:D5"), referenceRange:sheet.Range.get("E5")});
sparklines.Add({dataRange:sheet.Range.get("A6:D6"), referenceRange:sheet.Range.get("E6")});
sparklines.Add({dataRange:sheet.Range.get("A7:D7"), referenceRange:sheet.Range.get("E7")});
sparklines.Add({dataRange:sheet.Range.get("A8:D8"), referenceRange:sheet.Range.get("E8")});
sparklines.Add({dataRange:sheet.Range.get("A9:D9"), referenceRange:sheet.Range.get("E9")});
sparklines.Add({dataRange:sheet.Range.get("A10:D10"), referenceRange:sheet.Range.get("E10")});
sparklines.Add({dataRange:sheet.Range.get("A11:D11"), referenceRange:sheet.Range.get("E11")});
sparklines.Add({dataRange:sheet.Range.get("A2:D2"), referenceRange:sheet.Range.get("E2")});
sparklines.Add({dataRange:sheet.Range.get("A2:D2"), referenceRange:sheet.Range.get("E2")});
```

---

# spire.xls javascript chart
## create stacked column chart
```javascript
// Add a chart
let chart = sheet.Charts.Add();

// Set region of chart data
chart.DataRange = sheet.Range.get("A1:C5");
chart.SeriesDataFromRange = false;

// Set position of chart
chart.LeftColumn = 1;
chart.TopRow = 6;
chart.RightColumn = 11;
chart.BottomRow = 29;
chart.ChartType = wasmModule.ExcelChartType.ColumnStacked;

// Chart title
chart.ChartTitle = "Sales market by country";
chart.ChartTitleArea.IsBold = true;
chart.ChartTitleArea.Size = 12;

// Chart Axes
chart.PrimaryCategoryAxis.Title = "Country";
chart.PrimaryCategoryAxis.Font.IsBold = true;
chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
chart.PrimaryValueAxis.HasMajorGridLines = false;
chart.PrimaryValueAxis.MinValue = 1000;
chart.PrimaryValueAxis.TitleArea.IsBold = true;
chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;

for (let i = 0; i < chart.Series.Count; i++) {
    let cs = chart.Series.get(i);
    cs.Format.Options.IsVaryColor = true;
    cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
}

// Chart Legend
chart.Legend.Position = wasmModule.LegendPositionType.Top;
```

---

# spire.xls javascript shapes
## add arrow lines to excel file
```javascript
// Add a Double Arrow and fill the line with solid color.
let line = sheet.TypedLines.AddLine();
line.Top = 10;
line.Left = 20;
line.Width = 100;
line.Height = 0;
line.Color = wasmModule.Color.get_Blue();
line.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
line.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

// Add an Arrow and fill the line with solid color.
let line_1 = sheet.TypedLines.AddLine();
line_1.Top = 50;
line_1.Left = 30;
line_1.Width = 100;
line_1.Height = 100;
line_1.Color = wasmModule.Color.get_Red();
line_1.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineNoArrow;
line_1.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

// Add an Elbow Arrow Connector.
let line3 = sheet.TypedLines.AddLine();
line3.LineShapeType = wasmModule.LineShapeType.ElbowLine;
line3.Width = 30;
line3.Height = 50;
line3.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
line3.Top = 100;
line3.Left = 50;

// Add an Elbow Double-Arrow Connector.
let line2 = sheet.TypedLines.AddLine();
line2.LineShapeType = wasmModule.LineShapeType.ElbowLine;
line2.Width = 50;
line2.Height = 50;
line2.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
line2.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
line2.Left = 120;
line2.Top = 100;

// Add a Curved Arrow Connector.
line3 = sheet.TypedLines.AddLine();
line3.LineShapeType = wasmModule.LineShapeType.CurveLine;
line3.Width = 30;
line3.Height = 50;
line3.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowOpen;
line3.Top = 100;
line3.Left = 200;

// Add a Curved Double-Arrow Connector.
line2 = sheet.TypedLines.AddLine();
line2.LineShapeType = wasmModule.LineShapeType.CurveLine;
line2.Width = 30;
line2.Height = 50;
line2.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowOpen;
line2.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowOpen;
line2.Left = 250;
line2.Top = 100;
```

---

# spire.xls javascript shapes
## add line shapes to excel worksheet
```javascript
// Add shape line1
let line1 = sheet.Lines.AddLine({ row: 10, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.Line });
line1.DashStyle = wasmModule.ShapeDashLineStyleType.Solid;
line1.Color = wasmModule.Color.get_CadetBlue();
line1.Weight = 2;
line1.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

// Add shape line2
let line2 = sheet.Lines.AddLine({ row: 12, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.CurveLine });
line2.DashStyle = wasmModule.ShapeDashLineStyleType.Dotted;
line2.Color = wasmModule.Color.get_OrangeRed();
line2.Weight = 2;

// Add shape line3
let line3 = sheet.Lines.AddLine({ row: 14, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.ElbowLine });
line3.DashStyle = wasmModule.ShapeDashLineStyleType.DashDotDot;
line3.Color = wasmModule.Color.get_Purple();
line3.Weight = 2;

// Add shape line4
let line4 = sheet.Lines.AddLine({ row: 16, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.LineInv });
line4.DashStyle = wasmModule.ShapeDashLineStyleType.Dashed;
line4.Color = wasmModule.Color.get_Green();
line4.Weight = 2;
```

---

# Spire.XLS JavaScript Oval Shape
## Add oval shapes to Excel worksheet with different fill styles
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Add oval shape1
let ovalShape1 = sheet.OvalShapes.AddOval(11, 2, 100, 100);
ovalShape1.Line.Weight = 0;
// Fill shape with solid color
ovalShape1.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
ovalShape1.Fill.ForeColor = wasmModule.Color.get_DarkCyan();

// Add oval shape2
let ovalShape2 = sheet.OvalShapes.AddOval(11, 5, 100, 100);
ovalShape2.Line.Weight = 1;
// Fill shape with picture
ovalShape2.Line.DashStyle = wasmModule.ShapeDashLineStyleType.Solid;
ovalShape2.Fill.CustomPicture("logo.png");
```

---

# spire.xls javascript shapes
## add rectangle shapes to Excel worksheet
```javascript
// Add rectangle shape 1------Rect
let rect1 = sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, wasmModule.RectangleShapeType.Rect);
rect1.Line.Weight = 1;
// Fill shape with solid color
rect1.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
rect1.Fill.ForeColor = wasmModule.Color.get_DarkGreen();

// Add rectangle shape 2------RoundRect
let rect2 = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100,wasmModule.RectangleShapeType.RoundRect);
rect2.Line.Weight = 1;
rect2.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
rect2.Fill.ForeColor = wasmModule.Color.get_DarkCyan();
```

---

# Adding Spinner Control in Excel
## This code demonstrates how to add a spinner control to an Excel worksheet
```javascript
// Set text for range C11
sheet.Range.get("C11").Text = "Value:";
sheet.Range.get("C11").Style.Font.IsBold = true;

// Set value for range C12
sheet.Range.get("C12").Value2 = wasmModule.Int32.Create(0);

// Add spinner control
let spinner = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20);
spinner.LinkedCell = sheet.Range.get("C12");
spinner.Min = 0;
spinner.Max = 100;
spinner.IncrementalChange = 5;
spinner.Display3DShading = true;
```

---

# spire.xls javascript shapes
## adjust arrow polyline position in excel
```javascript
// Draw an elbow arrow
let line = worksheet.TypedLines.AddLine({
    row: 5,
    column: 5,
    width: 100,
    height: 100,
    lineShapeType: wasmModule.LineShapeType.ElbowLine
});
line.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineNoArrow;
line.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
let ad = line.ShapeAdjustValues.AddAdjustValue(wasmModule.GeomertyAdjustValueFormulaType.LiteralValue);

// When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
ad.SetFormulaParameter([-50]);
```

---

# Spire.XLS JavaScript Shape Copying
## Copy shapes between worksheets in Excel
```javascript
// Create line shape
let line = sheet.TypedLines.AddLine();
line.Top = 50;
line.Left = 30;
line.Width = 30;
line.Height = 50;
line.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowDiamond;
line.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

let copySheet = workbook.Worksheets.get(1);
// Copy the line into another sheet
copySheet.TypedLines.AddCopy(line);

// Create a button and then copy into another sheet
let button = sheet.TypedRadioButtons.Add({ row: 5, column: 5, height: 20, width: 20 });
copySheet.TypedRadioButtons.AddCopy(button);

// Create a textbox and then copy into another sheet
let textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100);
copySheet.TypedTextBoxes.AddCopy(textbox);

// Create a checkbox and then copy into another sheet
let checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20);
copySheet.TypedCheckBoxes.AddCopy(checkbox);

// Create a combobox and then copy into another sheet
sheet.Range.get("A14").Value = "1";
sheet.Range.get("A15").Value = "2";
let comboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30);
comboBoxes.ListFillRange = sheet.Range.get("A14:A15");
copySheet.TypedComboBoxes.AddCopy(comboBoxes);
```

---

# Delete All Shapes in Excel
## This code demonstrates how to delete all shapes from an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Delete all shapes in the worksheet
for (let i = sheet.PrstGeomShapes.Count - 1; i >= 0; i--) {
    sheet.PrstGeomShapes.get(i).Remove();
}
```

---

# Excel Shape Deletion
## Delete a particular shape from Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Delete the first shape in the worksheet
sheet.PrstGeomShapes.get(0).Remove();
```

---

# Excel Line Drawing
## Drawing lines through two points in Excel using JavaScript
```javascript
//1)Draw a line according to relative position
let line1 = worksheet.TypedLines.AddLine();
line1.LeftColumn = 3;
line1.TopRow = 3;
line1.LeftColumnOffset = 0;
line1.TopRowOffset = 0;

line1.RightColumn = 4;
line1.BottomRow = 5;
line1.RightColumnOffset = 0;
line1.BottomRowOffset = 0;

//2)Draw a line according to absolute position(pixels).
let line2 = worksheet.TypedLines.AddLine();
line2.StartPoint = wasmModule.Point.Create(30, 50);
line2.EndPoint = wasmModule.Point.Create(20, 80);
```

---

# Extract Text and Image from Excel Shapes
## This code demonstrates how to extract text and images from shapes in an Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Extract text from the first shape and save to a txt file
let shape1 = sheet.PrstGeomShapes.get(2);
let s = shape1.Text;
let sb = [];
sb.push(`The text in the third shape is: ${s}`);

// Extract image from shape
let shape2 = sheet.PrstGeomShapes.get(1);
let image = shape2.Fill.Picture;
let imageFile = `ExtractTextImageFromShape.png`;
image.Save(imageFile);
```

---

# Get Shape Linked Cell Range
## Retrieve the cell range addresses linked to Excel shapes
```javascript
// Load the workbook
workbook.LoadFromFile(excelFileName);
let sheet = workbook.Worksheets.get(0);
let prstGeomShapeCollection = sheet.PrstGeomShapes;

// Get the linked cell range address for a shape named "Yesterday"
let shape = prstGeomShapeCollection.get({name:"Yesterday"});
let cellAddress = shape.LinkedCell.RangeAddress;
sb.push(`${cellAddress}\n`);

// Get the linked cell range address for a shape named "NewShapes"
shape = prstGeomShapeCollection.get({name:"NewShapes"});
cellAddress = shape.LinkedCell.RangeAddress;
sb.push(cellAddress);
```

---

# Excel Shape Visibility Control
## Hide or unhide shapes in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Hide the second shape in the worksheet
sheet.PrstGeomShapes.get(1).Visible = false;

// Show the second shape in the worksheet
// sheet.PrstGeomShapes.get(1).Visible = true;
```

---

# spire.xls javascript shapes
## insert shapes to excel worksheet
```javascript
// Create a new workbook object
const workbook = wasmModule.Workbook.Create();
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Add a triangle shape.
let triangle = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, wasmModule.PrstGeomShapeType.Triangle);
//Fill the triangle with solid color.
triangle.Fill.ForeColor = wasmModule.Color.get_Yellow();
triangle.Fill.FillType = wasmModule.ShapeFillType.SolidColor;

//Add a heart shape.
let heart = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, wasmModule.PrstGeomShapeType.Heart);
//Fill the heart with gradient color.
heart.Fill.ForeColor = wasmModule.Color.get_Red();
heart.Fill.FillType = wasmModule.ShapeFillType.Gradient;

//Add an arrow shape with default color.
let arrow = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, wasmModule.PrstGeomShapeType.CurvedRightArrow);

//Add a cloud shape.
let cloud = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, wasmModule.PrstGeomShapeType.Cloud);
//Fill the cloud with custom picture
cloud.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile("wasmModule.png"), name:"wasmModule.png"});
cloud.Fill.FillType = wasmModule.ShapeFillType.Picture;
```

---

# spire.xls javascript shape
## modify shadow style for shape
```javascript
//Get the third shape from the worksheet.
let shape = sheet.PrstGeomShapes.get(2);

//Set the shadow style for the shape.
shape.Shadow.Angle = 90;
shape.Shadow.Transparency = 30;
shape.Shadow.Distance = 10;
shape.Shadow.Size = 130;
shape.Shadow.Color = wasmModule.Color.get_Yellow();
shape.Shadow.Blur = 30;
shape.Shadow.HasCustomStyle = true;
```

---

# Excel Shape Shadow Style Configuration
## Set shadow style properties for shapes in Excel worksheets
```javascript
// Create a new workbook object
const workbook = wasmModule.Workbook.Create();

//Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

//Add an ellipse shape.
let ellipse = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, wasmModule.PrstGeomShapeType.Ellipse);

//Set the shadow style for the ellipse.
ellipse.Shadow.Angle = 90;
ellipse.Shadow.Distance = 10;
ellipse.Shadow.Size = 150;
ellipse.Shadow.Color = wasmModule.Color.get_Gray();
ellipse.Shadow.Blur = 30;
ellipse.Shadow.Transparency = 1;
ellipse.Shadow.HasCustomStyle = true;
```

---

# spire.xls javascript shape order
## set the order of shapes in excel worksheets
```javascript
//Bring the picture forward one level
workbook.Worksheets.get(0).Pictures.get(0).ChangeLayer(wasmModule.ShapeLayerChangeType.BringForward);

//Bring the image in front of all other objects
workbook.Worksheets.get(1).Pictures.get(0).ChangeLayer(wasmModule.ShapeLayerChangeType.BringToFront);

//Send the shape back one level
let shape = workbook.Worksheets.get(2).PrstGeomShapes.get(1);
shape.ChangeLayer(wasmModule.ShapeLayerChangeType.SendBackward);

//Send the shape behind all other objects
shape = workbook.Worksheets.get(3).PrstGeomShapes.get(1);
shape.ChangeLayer(wasmModule.ShapeLayerChangeType.SendToBack);
```

---

# Shape to Image Conversion
## Convert Excel shape to image file
```javascript
// Get the first worksheet
let sheet1 = workbook.Worksheets.get(0);

// Get the first shape from the first worksheet
let shape = sheet1.PrstGeomShapes.get(0);

// Save the shape to an image
let img = shape.SaveToImage();
let outputFile = "ShapeToImage.png"
img.Save(outputFile);
```

---

# Spire.XLS JavaScript Shape Texture
## Fill a shape with tiled picture as texture
```javascript
//Get the first shape
let shape = sheet.PrstGeomShapes.get(0);

//Fill shape with texture
shape.Fill.FillType = wasmModule.ShapeFillType.Texture;

//Custom texture with picture
shape.Fill.CustomTexture({path:"Logo.png"});

//Tile picture as texture
shape.Fill.Tile = true;
```

---

# Spire.XLS JavaScript Styles
## Apply built-in styles to Excel range
```javascript
// Get the first sheet
const sheet = workbook.Worksheets.get(0);

// Apply title style
sheet.Range.get("A1:J1").BuiltInStyle = wasmModule.BuiltInStyles.Title;
```

---

# Excel Color Scales Formatting
## Apply color scales to data range in Excel
```javascript
//Create a workbook.
const workbook = wasmModule.Workbook.Create();

//Get the first worksheet.
const sheet = workbook.Worksheets.get(0);

//Add color scales.
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
const format = xcfs.AddCondition();
format.FormatType = wasmModule.ConditionalFormatType.ColorScale;
```

---

# spire.xls javascript conditional formatting
## apply conditional formatting to cell range in excel
```javascript
//Create conditional formatting rule.
const xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(sheet.AllocatedRange);
const format1 = xcfs1.AddCondition();
format1.FormatType = wasmModule.ConditionalFormatType.CellValue;
format1.FirstFormula = "800";
format1.Operator = wasmModule.ComparisonOperatorType.Greater;
format1.FontColor = wasmModule.Color.get_Red();
format1.BackColor = wasmModule.Color.get_LightSalmon();

//Create conditional formatting rule.
const xcfs2 = sheet.ConditionalFormats.Add();
xcfs2.AddRange(sheet.AllocatedRange);
const format2 = xcfs1.AddCondition();
format2.FormatType = wasmModule.ConditionalFormatType.CellValue;
format2.FirstFormula = "300";
format2.Operator = wasmModule.ComparisonOperatorType.Less;
format2.FontColor = wasmModule.Color.get_Green();
format2.BackColor = wasmModule.Color.get_LightBlue();
```

---

# Excel Data Bars Formatting
## Apply conditional formatting data bars to cell range
```javascript
//Create a workbook.
const workbook = wasmModule.Workbook.Create();

//Get the first worksheet.
const sheet = workbook.Worksheets.get(0);

//Insert data to cell range from A1 to C4.
sheet.Range.get("A1").NumberValue = 582;
sheet.Range.get("A2").NumberValue = 234;
sheet.Range.get("A3").NumberValue = 314;
sheet.Range.get("A4").NumberValue = 50;
sheet.Range.get("B1").NumberValue = 150;
sheet.Range.get("B2").NumberValue = 894;
sheet.Range.get("B3").NumberValue = 560;
sheet.Range.get("B4").NumberValue = 900;
sheet.Range.get("C1").NumberValue = 134;
sheet.Range.get("C2").NumberValue = 700;
sheet.Range.get("C3").NumberValue = 920;
sheet.Range.get("C4").NumberValue = 450;
sheet.AllocatedRange.RowHeight = 15;
sheet.AllocatedRange.ColumnWidth = 17;

//Add data bars.
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
const format = xcfs.AddCondition();
format.FormatType = wasmModule.ConditionalFormatType.DataBar;
format.DataBar.BarColor = wasmModule.Color.get_CadetBlue();
```

---

# Excel Gradient Fill Effects
## Apply gradient filling effects to Excel cells
```javascript
//Get "B5" cell
const range = sheet.Range.get("B5");

//Set gradient filling effects
range.Style.Interior.FillPattern = wasmModule.ExcelPatternType.Gradient;
range.Style.Interior.Gradient.ForeColor = wasmModule.Color.FromArgb(255, 255, 255);
range.Style.Interior.Gradient.BackColor = wasmModule.Color.FromArgb(79, 129, 189);
range.Style.Interior.Gradient.TwoColorGradient(wasmModule.GradientVariantsType.HorizontalAlignTypeHorizontal, wasmModule.GradientVariantsType.Color);
```

---

# Excel Icon Sets Application
## Apply icon sets to cell range in Excel
```javascript
//Add icon sets.
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
const format = xcfs.AddCondition();
format.FormatType = wasmModule.ConditionalFormatType.IconSet;
format.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights1;
```

---

# Excel Colors and Palette Management
## Setting custom colors and palette for Excel workbook
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Adding Orchid color to the palette at 60th index
workbook.ChangePaletteColor(wasmModule.Color.get_Orchid(), 60);

// Get the first sheet
const sheet = workbook.Worksheets.get(0);

const cell = sheet.Range.get("B2");
cell.Text = "Welcome to use Spire.XLS";

// Set the Orchid (custom) color to the font
cell.Style.Font.Color = wasmModule.Color.get_Orchid();
cell.Style.Font.Size = 20;
cell.AutoFitColumns();
cell.AutoFitRows();
```

---

# spire.xls javascript conditional formatting
## add runtime conditional formatting rules to excel cells
```javascript
function AddComparisonRule1(sheet) {
    //Create conditional formatting rule
    const xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range.get("A1:D1"));
    const cf1 = xcfs1.AddCondition();
    cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
    cf1.FirstFormula = "150";
    cf1.Operator = wasmModule.ComparisonOperatorType.Greater;
    cf1.FontColor = wasmModule.Color.get_Red();
    cf1.BackColor = wasmModule.Color.get_LightBlue();
}

function AddComparisonRule2(sheet) {
    const xcfs2 = sheet.ConditionalFormats.Add();
    xcfs2.AddRange(sheet.Range.get("A2:D2"));
    const cf2 = xcfs2.AddCondition();
    cf2.FormatType = wasmModule.ConditionalFormatType.CellValue;
    cf2.FirstFormula = "500";
    cf2.Operator = wasmModule.ComparisonOperatorType.Less;
    //Set border color
    cf2.LeftBorderColor = wasmModule.Color.get_Pink();
    cf2.RightBorderColor = wasmModule.Color.get_Pink();
    cf2.TopBorderColor = wasmModule.Color.get_DeepSkyBlue();
    cf2.BottomBorderColor = wasmModule.Color.get_DeepSkyBlue();
    cf2.LeftBorderStyle = wasmModule.LineStyleType.Medium;
    cf2.RightBorderStyle = wasmModule.LineStyleType.Thick;
    cf2.TopBorderStyle = wasmModule.LineStyleType.Double;
    cf2.BottomBorderStyle = wasmModule.LineStyleType.Double;
}

function AddComparisonRule3(sheet) {
    //Create conditional formatting rule
    const xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range.get("A3:D3"));
    const cf1 = xcfs1.AddCondition();
    cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
    cf1.FirstFormula = "300";
    cf1.SecondFormula = "500";
    cf1.Operator = wasmModule.ComparisonOperatorType.Between;
    cf1.BackColor = wasmModule.Color.get_Yellow();
}

function AddComparisonRule4(sheet) {
    //Create conditional formatting rule
    const xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range.get("A4:D4"));
    const cf1 = xcfs1.AddCondition();
    cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
    cf1.FirstFormula = "100";
    cf1.SecondFormula = "200";
    cf1.Operator = wasmModule.ComparisonOperatorType.NotBetween;
    //Set fill pattern type
    cf1.FillPattern = wasmModule.ExcelPatternType.ReverseDiagonalStripe;
    //Set foreground color
    cf1.Color = wasmModule.Color.FromArgb(255, 255, 0);
    //Set background color
    cf1.BackColor = wasmModule.Color.FromArgb(0, 255, 255);
}

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

AddComparisonRule1(sheet);
AddComparisonRule2(sheet);
AddComparisonRule3(sheet);
AddComparisonRule4(sheet);
```

---

# Excel Conditional Date Formatting
## Apply conditional formatting to highlight dates in the last 7 days
```javascript
// Get the first worksheet
const sheet = workbook.Worksheets.get(0);

// Highlight cells that contain a date occurring in the last 7 days
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.AllocatedRange);
const conditionalFormat = xcfs.AddTimePeriodCondition(wasmModule.TimePeriodType.Last7Days);
conditionalFormat.BackColor = wasmModule.Color.get_Orange();
```

---

# Excel Formula Conditional Formatting
## Create formula-based conditional formatting rule in Excel
```javascript
// Get the first worksheet and the first column from the workbook
const sheet = workbook.Worksheets.get(0);
const range = sheet.Columns.get(0);

// Set the conditional formatting formula and apply the rule to the chosen cell range
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(range);
const conditional = xcfs.AddCondition();
conditional.FormatType = wasmModule.ConditionalFormatType.Formula;
conditional.FirstFormula = "=($A1<$B1)";
conditional.BackKnownColor = wasmModule.ExcelColors.Yellow;
```

---

# Excel Font Styles Formatting
## Apply various font styles to Excel cells
```javascript
// Get the first sheet
const sheet = workbook.Worksheets.get(0);

// Set font style
sheet.Range.get("B1").Style.Font.FontName = "Comic Sans MS";
sheet.Range.get("B2:D2").Style.Font.FontName = "Corbel";
sheet.Range.get("B3:D7").Style.Font.FontName = "Aleo";

// Set font size
sheet.Range.get("B1").Style.Font.Size = 45;
sheet.Range.get("B2:D3").Style.Font.Size = 25;
sheet.Range.get("B3:D7").Style.Font.Size = 12;

// Set excel cell data to be bold
sheet.Range.get("B2:D2").Style.Font.IsBold = true;

// Set excel cell data to be underline
sheet.Range.get("B3:B7").Style.Font.Underline = wasmModule.FontUnderlineType.Single;

// Set excel cell data color
sheet.Range.get("B1").Style.Font.Color = wasmModule.Color.get_CornflowerBlue();
sheet.Range.get("B2:D2").Style.Font.Color = wasmModule.Color.get_CadetBlue();
sheet.Range.get("B3:D7").Style.Font.Color = wasmModule.Color.get_Firebrick();

// Set excel cell data to be italic
sheet.Range.get("B3:D7").Style.Font.IsItalic = true;

// Add strikethrough
sheet.Range.get("D3").Style.Font.IsStrikethrough = true;
sheet.Range.get("D7").Style.Font.IsStrikethrough = true;
```

---

# Excel Cell Formatting
## Set foreground and background colors for Excel cells
```javascript
//Create a new style
const style1 = workbook.Styles.Add("newStyle1");

//Set filling pattern type
style1.Interior.FillPattern = wasmModule.ExcelPatternType.Gradient;

//Set filling Background color
style1.Interior.Gradient.BackKnownColor = wasmModule.ExcelColors.Green;

//Set filling Foreground color
style1.Interior.Gradient.ForeKnownColor = wasmModule.ExcelColors.Yellow;

//Apply the style to "B2" cell
sheet.Range.get("B2").CellStyleName = style1.Name;
sheet.Range.get("B2").Text = "Test";
sheet.Range.get("B2").RowHeight = 30;
sheet.Range.get("B2").ColumnWidth = 50;

//Create a new style
const style2 = workbook.Styles.Add("newStyle2");

//Set filling pattern type
style2.Interior.FillPattern = wasmModule.ExcelPatternType.Gradient;

//Set filling Foreground color
style2.Interior.Gradient.ForeKnownColor = wasmModule.ExcelColors.Red;

//Apply the style to "B4" cell
sheet.Range.get("B4").CellStyleName = style2.Name;
sheet.Range.get("B4").RowHeight = 30;
sheet.Range.get("B4").ColumnWidth = 60;
```

---

# Excel Column Formatting
## Format a column in Excel with custom style properties
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

//Create a new style
const style = workbook.Styles.Add("newStyle");

//Set the vertical alignment of the text
style.VerticalAlignment = wasmModule.VerticalAlignType.Center;

//Set the horizontal alignment of the text
style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

//Set the font color of the text
style.Font.Color = wasmModule.Color.get_Blue();

//Shrink the text to fit in the cell
style.ShrinkToFit = true;

//Set the bottom border color of the cell to OrangeRed
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).Color = wasmModule.Color.get_OrangeRed();

//Set the bottom border type of the cell to Dotted
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle = wasmModule.LineStyleType.Dotted;

//Apply the style to the first column
sheet.Columns.get(0).CellStyleName = style.Name;

sheet.Columns.get(0).Text = "Test";
```

---

# Excel Row Formatting
## Apply style formatting to a specific row in Excel
```javascript
//Create a new style
const style = workbook.Styles.Add("newStyle");

//Set the vertical alignment of the text
style.VerticalAlignment = wasmModule.VerticalAlignType.Center;

//Set the horizontal alignment of the text
style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

//Set the font color of the text
style.Font.Color = wasmModule.Color.get_Blue();

//Shrink the text to fit in the cell
style.ShrinkToFit = true;

//Set the bottom border color of the cell to OrangeRed
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).Color = wasmModule.Color.get_OrangeRed();

//Set the bottom border type of the cell to Dotted
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle = wasmModule.LineStyleType.Dotted;

//Apply the style to the second row
sheet.Rows.get(1).CellStyleName = style.Name;

sheet.Rows.get(1).Text = "Test";
```

---

# Excel Cell Style Formatting
## Create and apply a custom style to cells in an Excel worksheet
```javascript
//Create a style
const style = workbook.Styles.Add("newStyle");
//Set the shading color
style.Color = wasmModule.Color.get_DarkGray();
//Set the font color
style.Font.Color = wasmModule.Color.get_White();
//Set font name
style.Font.FontName = "Times New Roman";
//Set font size
style.Font.Size = 12;
//Set bold for the font
style.Font.IsBold = true;
//Set text rotation
style.Rotation = 45;
//Set alignment
style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;
style.VerticalAlignment = wasmModule.VerticalAlignType.Center;

//Set the style for the specific range
workbook.Worksheets.get(0).Range.get("A1:J1").CellStyleName = style.Name;
```

---

# spire.xls javascript color
## get color ARGB data from Excel cells
```javascript
//Get the first sheet
const sheet = workbook.Worksheets.get(0);

const strB = [];

//Get font color
const color1 = sheet.Range.get("B2").Style.Font.Color;

//Read ARGB data of Color
strB.push(`The font color of B2: ARGB=(${color1.A},${color1.R},${color1.G},${color1.B})`);

const color2 = sheet.Range.get("B3").Style.Font.Color;
strB.push(`The font color of B3: ARGB=(${color2.A},${color2.R},${color2.G},${color2.B})`);

const color3 = sheet.Range.get("B4").Style.Font.Color;
strB.push(`The font color of B4: ARGB=(${color3.A},${color3.R},${color3.G},${color3.B})`);
```

---

# spire.xls javascript style
## get and set cell style
```javascript
//Get the first sheet
const sheet = workbook.Worksheets.get(0);

//Get "B4" cell
const range = sheet.Range.get("B4");
//Get the style of cell
const style = range.Style;
style.Font.FontName = "Calibri";
style.Font.IsBold = true;
style.Font.Size = 15;
style.Font.Color = wasmModule.Color.get_CornflowerBlue();
range.Style = style;
```

---

# spire.xls javascript conditional formatting
## highlight above and below average values in Excel
```javascript
//Add conditional format.
const format1 = sheet.ConditionalFormats.Add();
//Set the cell range to apply the formatting.
format1.AddRange(sheet.Range.get("E2:E10"));
//Add below average condition.
const cf1 = format1.AddAverageCondition(wasmModule.AverageType.Below);
//Highlight cells below average values.
cf1.BackColor = wasmModule.Color.get_SkyBlue();

//Add conditional format.
const format2 = sheet.ConditionalFormats.Add();
//Set the cell range to apply the formatting.
format2.AddRange(sheet.Range.get("E2:E10"));
//Add above average condition.
const cf2 = format2.AddAverageCondition(wasmModule.AverageType.Above);
//Highlight cells above average values.
cf2.BackColor = wasmModule.Color.get_Orange();
```

---

# Spire.XLS JavaScript Conditional Formatting
## Highlight duplicate and unique values in Excel
```javascript
//Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.Range.get("C2:C10"));
const format1 = xcfs.AddCondition();
format1.FormatType = wasmModule.ConditionalFormatType.DuplicateValues;
format1.BackColor = wasmModule.Color.get_IndianRed();

//Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
const xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(sheet.Range.get("C2:C10"));
const format2 = xcfs1.AddCondition();
format2.FormatType = wasmModule.ConditionalFormatType.UniqueValues;
format2.BackColor = wasmModule.Color.get_Yellow();
```

---

# Excel Conditional Formatting
## Highlight top and bottom ranked values in Excel cells
```javascript
//Get the first worksheet.
const sheet = workbook.Worksheets.get(0);

//Apply conditional formatting to range "D2:D10" to highlight the top 2 values.
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(sheet.Range.get("D2:D10"));
const format1 = xcfs.AddTopBottomCondition(wasmModule.TopBottomType.Top, 2);
format1.FormatType = wasmModule.ConditionalFormatType.TopBottom;
format1.BackColor = wasmModule.Color.get_Red();

//Apply conditional formatting to range "E2:E10" to highlight the bottom 2 values.
const xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(sheet.Range.get("E2:E10"));
const format2 = xcfs1.AddTopBottomCondition(wasmModule.TopBottomType.Bottom, 2);
format2.FormatType = wasmModule.ConditionalFormatType.TopBottom;
format2.BackColor = wasmModule.Color.get_ForestGreen();
```

---

# Excel Cell Indentation
## Set text indentation level in Excel cells
```javascript
//Access the "B5" cell from the worksheet
const cell = sheet.Range.get("B5");

//Add some value to the "B5" cell
cell.Text = "Hello Spire!";

//Set the indentation level of the text (inside the cell) to 2
cell.Style.IndentLevel = 2;
```

---

# Excel Cell Interior Styling
## Set gradient interior styles for Excel cells
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Initialize the workbook
const sheet = workbook.Worksheets.get(0);

//Specify the version
workbook.Version = wasmModule.ExcelVersion.Version2007;

//Define the number of the colors
const maxColor = Object.keys(wasmModule.ExcelColors).length;

//Create a random object
for (let i = 2; i < 40; i++)
{
    //Random backKnownColor
    const backKnownColor = wasmModule.ExcelColors.fromValue(Math.floor(Math.random() * (maxColor / 2)));

    //Add text
    sheet.Range.get("A1").Text = "Color Name";
    sheet.Range.get("B1").Text = "Red";
    sheet.Range.get("C1").Text = "Green";
    sheet.Range.get("D1").Text = "Blue";

    //Merge the sheet"E1-K1"
    sheet.Range.get("E1:K1").Merge();
    sheet.Range.get("E1:K1").Text = "Gradient";
    sheet.Range.get("A1:K1").Style.Font.IsBold = true;
    sheet.Range.get("A1:K1").Style.Font.Size = 11;

    //Set the text of color in sheetA-sheetD
    const colorName = backKnownColor;
    sheet.Range.get(`A${i}`).Text = colorName;
    sheet.Range.get(`B${i}`).NumberValue = workbook.GetPaletteColor(backKnownColor).R;
    sheet.Range.get(`C${i}`).NumberValue = workbook.GetPaletteColor(backKnownColor).G;
    sheet.Range.get(`D${i}`).NumberValue = workbook.GetPaletteColor(backKnownColor).B;

    //Merge the sheets
    sheet.Range.get(`E${i}:K${i}`).Merge();

    //Set the text of sheetE-sheetK
    sheet.Range.get(`E${i}:K${i}`).Text = colorName;

    //Set the interior of the color
    sheet.Range.get(`E${i}:K${i}`).Style.Interior.FillPattern = wasmModule.ExcelPatternType.Gradient;
    sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.BackKnownColor = backKnownColor;
    sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.ForeKnownColor = wasmModule.ExcelColors.White;
    sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.GradientStyle = wasmModule.GradientStyleType.Vertical;
    sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.GradientVariant = wasmModule.GradientVariantsType.ShadingVariants1;
}

//AutoFit Column
sheet.AutoFitColumn(1);
```

---

# Excel Cell Activation
## Make a specific cell active in an Excel worksheet
```javascript
//Get the 2nd sheet
const sheet = workbook.Worksheets.get(1);

//Set the 2nd sheet as an active sheet
sheet.Activate();

//Set B2 cell as an active cell in the worksheet
sheet.SetActiveCell(sheet.Range.get("B2"));

//Set the B column as the first visible column in the worksheet
sheet.FirstVisibleColumn = 1;

//Set the 2nd row as the first visible row in the worksheet
sheet.FirstVisibleRow = 1;
```

---

# Excel Number Formatting
## Demonstrates various number formatting options in Excel cells
```javascript
// Input a number value for the specified cell and set the number format
sheet.Range.get("B10").Text = "NUMBER FORMATTING";
sheet.Range.get("B10").Style.Font.IsBold = true;

sheet.Range.get("B13").Text = "0";
sheet.Range.get("C13").NumberValue = 1234.5678;
sheet.Range.get("C13").NumberFormat = "0";

sheet.Range.get("B14").Text = "0.00";
sheet.Range.get("C14").NumberValue = 1234.5678;
sheet.Range.get("C14").NumberFormat = "0.00";

sheet.Range.get("B15").Text = "#,##0.00";
sheet.Range.get("C15").NumberValue = 1234.5678;
sheet.Range.get("C15").NumberFormat = "#,##0.00";

sheet.Range.get("B16").Text = "$#,##0.00";
sheet.Range.get("C16").NumberValue = 1234.5678;
sheet.Range.get("C16").NumberFormat = "$#,##0.00";

sheet.Range.get("B17").Text = "0;[Red]-0";
sheet.Range.get("C17").NumberValue = -1234.5678;
sheet.Range.get("C17").NumberFormat = "0;[Red]-0";

sheet.Range.get("B18").Text = "0.00;[Red]-0.00";
sheet.Range.get("C18").NumberValue = -1234.5678;
sheet.Range.get("C18").NumberFormat = "0.00;[Red]-0.00";

sheet.Range.get("B19").Text = "#,##0;[Red]-#,##0";
sheet.Range.get("C19").NumberValue = -1234.5678;
sheet.Range.get("C19").NumberFormat = "#,##0;[Red]-#,##0";

sheet.Range.get("B20").Text = "#,##0.00;[Red]-#,##0.00";
sheet.Range.get("C20").NumberValue = -1234.5678;
sheet.Range.get("C20").NumberFormat = "#,##0.00;[Red]-#,##0.00";

sheet.Range.get("B21").Text = "0.00E+00";
sheet.Range.get("C21").NumberValue = 1234.5678;
sheet.Range.get("C21").NumberFormat = "0.00E+00";

sheet.Range.get("B22").Text = "0.00%";
sheet.Range.get("C22").NumberValue = 1234.5678;
sheet.Range.get("C22").NumberFormat = "0.00%";

sheet.Range.get("B13:B22").Style.KnownColor = wasmModule.ExcelColors.Gray25Percent;

// AutoFit Column
sheet.AutoFitColumn(2);
sheet.AutoFitColumn(3);
```

---

# Spire.XLS JavaScript Border Formatting
## Set cell border styles in Excel worksheets
```javascript
//Get the first worksheet
const sheet = workbook.Worksheets.get(0);

//Get the cell range where you want to apply border style
const cr = sheet.Range.get({
  row: sheet.FirstRow,
  column: sheet.FirstColumn,
  lastRow: sheet.LastRow,
  lastColumn: sheet.LastColumn
});

//Apply border style
cr.Borders.LineStyle = wasmModule.LineStyleType.Double;
cr.Borders.get(wasmModule.BordersLineType.DiagonalDown).LineStyle = wasmModule.LineStyleType.None;
cr.Borders.get(wasmModule.BordersLineType.DiagonalUp).LineStyle = wasmModule.LineStyleType.None;
cr.Borders.Color = wasmModule.Color.get_CadetBlue();
```

---

# Excel DataBar Border Formatting
## Set border to DataBar in Excel using JavaScript
```javascript
//Get the databar format 
const xcfs = sheet.ConditionalFormats.get(0);
const cf = xcfs.get(0);
const dataBar1 = cf.DataBar;
dataBar1.BarBorder.Type = wasmModule.DataBarBorderType.DataBarBorderSolid;
dataBar1.BarBorder.Color = wasmModule.Color.get_Red();

//Set to new data bar
sheet.Range.get("E1").NumberValue = 200;
const xcfs2 = sheet.ConditionalFormats.Add();
xcfs2.AddRange(sheet.Range.get("E1"));
const cf2 = xcfs2.AddCondition();
cf2.FormatType = wasmModule.ConditionalFormatType.DataBar;
cf2.DataBar.BarBorder.Type = wasmModule.DataBarBorderType.DataBarBorderSolid;
cf2.DataBar.BarBorder.Color = wasmModule.Color.get_Red();
cf2.DataBar.BarColor = wasmModule.Color.get_GreenYellow();
```

---

# spire.xls javascript conditional formatting
## set conditional format formula in worksheet
```javascript
//Add ConditionalFormat
const xcfs = sheet.ConditionalFormats.Add();

//Define the range
xcfs.AddRange(sheet.Range.get("B5"));

//Add condition
const format = xcfs.AddCondition();
format.FormatType = wasmModule.ConditionalFormatType.CellValue;

//If greater than 1000
format.FirstFormula = "1000";
format.Operator = wasmModule.ComparisonOperatorType.Greater;
format.BackColor = wasmModule.Color.get_Orange();

//Set a SUM formula for B5
sheet.get("B5").Formula = "=SUM(B1:B4)";

//Add text
sheet.get("C5").Text = "If Sum of B1:B4 is greater than 1000, B5 will have orange background.";
```

---

# Excel Row Conditional Formatting
## Set row colors based on even/odd rows using conditional formatting
```javascript
//Get the first worksheet
const sheet = workbook.Worksheets.get(0);

//Select the range that you want to format
const dataRange = sheet.AllocatedRange;

//Set conditional formatting
const xcfs = sheet.ConditionalFormats.Add();
xcfs.AddRange(dataRange);

const format1 = xcfs.AddCondition();
//Determines the cells to format
format1.FirstFormula = "=MOD(ROW(),2)=0";
//Set conditional formatting type
format1.FormatType = wasmModule.ConditionalFormatType.Formula;
//Set the color
format1.BackColor = wasmModule.Color.get_LightSeaGreen();

//Set the backcolor of the odd rows as Yellow
const xcfs1 = sheet.ConditionalFormats.Add();
xcfs1.AddRange(dataRange);

const format2 = xcfs1.AddCondition();
format2.FirstFormula = "=MOD(ROW(),2)=1";
format2.FormatType = wasmModule.ConditionalFormatType.Formula;
format2.BackColor = wasmModule.Color.get_Yellow();
```

---

# Excel Traffic Lights Icons Formatting
## Set traffic lights icons in Excel cells using conditional formatting
```javascript
// Add a conditional formatting
const conditional = sheet.ConditionalFormats.Add();
conditional.AddRange(sheet.AllocatedRange);
const format1 = conditional.AddCondition();

// Add a conditional formatting of cell range and set its type to CellValue
format1.FormatType = wasmModule.ConditionalFormatType.CellValue;
format1.FirstFormula = "300";
format1.Operator = wasmModule.ComparisonOperatorType.Less;
format1.FontColor = wasmModule.Color.get_Black();
format1.BackColor = wasmModule.Color.get_LightSkyBlue();

// Add a conditional formatting of cell range and set its type to IconSet
conditional.AddRange(sheet.AllocatedRange);
const format = conditional.AddCondition();
format.FormatType = wasmModule.ConditionalFormatType.IconSet;
format.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights1;
```

---

# spire.xls javascript conditional formatting
## apply various conditional formatting rules to excel cells
```javascript
function AddConditionalFormattingForExistingSheet(sheet) {
    sheet.AllocatedRange.RowHeight = 15;
    sheet.AllocatedRange.ColumnWidth = 16;

    // Create conditional formatting rule
    const xcfs1 = sheet.ConditionalFormats.Add();
    xcfs1.AddRange(sheet.Range.get("A1:D1"));
    const cf1 = xcfs1.AddCondition();
    cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
    cf1.FirstFormula = "150";
    cf1.Operator = wasmModule.ComparisonOperatorType.Greater;
    cf1.FontColor = wasmModule.Color.get_Red();
    cf1.BackColor = wasmModule.Color.get_LightBlue();

    const xcfs2 = sheet.ConditionalFormats.Add();
    xcfs2.AddRange(sheet.Range.get("A2:D2"));
    const cf2 = xcfs2.AddCondition();
    cf2.FormatType = wasmModule.ConditionalFormatType.CellValue;
    cf2.FirstFormula = "300";
    cf2.Operator = wasmModule.ComparisonOperatorType.Less;
    // Set border color
    cf2.LeftBorderColor = wasmModule.Color.get_Pink();
    cf2.RightBorderColor = wasmModule.Color.get_Pink();
    cf2.TopBorderColor = wasmModule.Color.get_DeepSkyBlue();
    cf2.BottomBorderColor = wasmModule.Color.get_DeepSkyBlue();
    cf2.LeftBorderStyle = wasmModule.LineStyleType.Medium;
    cf2.RightBorderStyle = wasmModule.LineStyleType.Thick;
    cf2.TopBorderStyle = wasmModule.LineStyleType.Double;
    cf2.BottomBorderStyle = wasmModule.LineStyleType.Double;

    // Add data bars
    const xcfs3 = sheet.ConditionalFormats.Add();
    xcfs3.AddRange(sheet.Range.get("A3:D3"));
    const cf3 = xcfs3.AddCondition();
    cf3.FormatType = wasmModule.ConditionalFormatType.DataBar;
    cf3.DataBar.BarColor = wasmModule.Color.get_CadetBlue();

    // Add icon sets
    const xcfs4 = sheet.ConditionalFormats.Add();
    xcfs4.AddRange(sheet.Range.get("A4:D4"));
    const cf4 = xcfs4.AddCondition();
    cf4.FormatType = wasmModule.ConditionalFormatType.IconSet;
    cf4.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights1;

    // Add color scales
    const xcfs5 = sheet.ConditionalFormats.Add();
    xcfs5.AddRange(sheet.Range.get("A5:D5"));
    const cf5 = xcfs5.AddCondition();
    cf5.FormatType = wasmModule.ConditionalFormatType.ColorScale;

    // Highlight duplicate values in range "A6:D6" with BurlyWood color
    const xcfs6 = sheet.ConditionalFormats.Add();
    xcfs6.AddRange(sheet.Range.get("A6:D6"));
    const cf6 = xcfs6.AddCondition();
    cf6.FormatType = wasmModule.ConditionalFormatType.DuplicateValues;
    cf6.BackColor = wasmModule.Color.get_BurlyWood();
}
```

---

# Excel Text Alignment Formatting
## Set vertical and horizontal alignment of text in Excel cells along with rotation
```javascript
// Set the vertical alignment to Top
sheet.Range.get("B1:C1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Top;

// Set the vertical alignment to Center
sheet.Range.get("B2:C2").Style.VerticalAlignment = wasmModule.VerticalAlignType.Center;

// Set the vertical alignment to Bottom
sheet.Range.get("B3:C3").Style.VerticalAlignment = wasmModule.VerticalAlignType.Bottom;

// Set the horizontal alignment to General
sheet.Range.get("B4:C4").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.General;

// Set the horizontal alignment to Left
sheet.Range.get("B5:C5").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Left;

// Set the horizontal alignment to Center
sheet.Range.get("B6:C6").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

// Set the horizontal alignment to Right
sheet.Range.get("B7:C7").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Right;

// Set the rotation degree
sheet.Range.get("B8:C8").Style.Rotation = 45;
sheet.Range.get("B9:C9").Style.Rotation = 90;

// Set the row height of cell
sheet.Range.get("B8:C9").RowHeight = 60;
```

---

# spire.xls javascript text direction
## set text direction from right to left in Excel cells
```javascript
// Access the "B5" cell from the worksheet
const cell = sheet.Range.get("B5");

// Add some value to the "B5" cell
cell.Text = "Hello Spire!";

// Set the reading order from right to left of the text in the "B5" cell
cell.Style.ReadingOrder = wasmModule.ReadingOrderType.RightToLeft;
```

---

# Excel Predefined Styles
## Create and apply predefined styles to Excel cells
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get the first sheet
const sheet = workbook.Worksheets.get(0);

// Create a new style
const style = workbook.Styles.Add("newStyle");
style.Font.FontName = "Calibri";
style.Font.IsBold = true;
style.Font.Size = 15;
style.Font.Color = wasmModule.Color.get_CornflowerBlue();

// Get "B5" cell
const range = sheet.Range.get("B5");
range.Text = "Welcome to use Spire.XLS";
range.CellStyleName = style.Name;
range.AutoFitColumns();
```

---

# Spire.XLS JavaScript Style Object
## Create and apply style objects to Excel cells
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Add a new worksheet to the Excel object
const sheet = workbook.Worksheets.Add("new sheet");

// Access the "B1" cell from the worksheet
const cell = sheet.Range.get("B1");

// Add some value to the "B1" cell
cell.Text = "Hello Spire!";

// Create a new style
const style = workbook.Styles.Add("newStyle");

// Set the vertical alignment of the text in the "B1" cell
style.VerticalAlignment = wasmModule.VerticalAlignType.Center;

// Set the horizontal alignment of the text in the "B1" cell
style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

// Set the font color of the text in the "B1" cell
style.Font.Color = wasmModule.Color.get_Blue();

// Shrink the text to fit in the cell
style.ShrinkToFit = true;

// Set the bottom border color of the cell to GreenYellow
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).Color = wasmModule.Color.get_GreenYellow();

// Set the bottom border type of the cell to Medium
style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle = wasmModule.LineStyleType.Medium;

// Assign the Style object to the "B1" cell
cell.Style = style;

// Apply the same style to some other cells
sheet.Range.get("B4").Style = style;
sheet.Range.get("B4").Text = "Test";
sheet.Range.get("C3").CellStyleName = style.Name;
sheet.Range.get("C3").Text = "Welcome to use Spire.XLS";
sheet.Range.get("D4").Style = style;
```

---

# Excel Conditional Formatting Implementation
## Various types of conditional formatting for Excel cells
```javascript
// Add default icon set conditional formatting
function AddDefaultIconSet(sheet) {
    const xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range.get("A1:C2"));
    sheet.Range.get("A1:C2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("A1:C2").Style.Color = wasmModule.Color.get_Yellow();
    const cf = xcfs.AddCondition();
    cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
    sheet.Range.get("A1").NumberValue = 0;
    sheet.Range.get("B1").NumberValue = 3;
    sheet.Range.get("C1").NumberValue = 6;
    sheet.Range.get("A2").NumberValue = 2;
    sheet.Range.get("B2").NumberValue = 5;
    sheet.Range.get("C2").NumberValue = 8;
}

// Add three arrows icon set conditional formatting
function AddIconSet2(sheet) {
    let xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range.get("M1:O2"));
    sheet.Range.get("M1:O2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("M1:O2").Style.Color = wasmModule.Color.get_AliceBlue();
    let cf = xcfs.AddCondition();
    cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
    cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeArrows;
    sheet.Range.get("M1").Text = "ThreeArrows";
    sheet.Range.get("N1").NumberValue = 15;
    sheet.Range.get("O1").NumberValue = 18;
    sheet.Range.get("M2").NumberValue = 14;
    sheet.Range.get("N2").NumberValue = 17;
    sheet.Range.get("O2").NumberValue = 20;
}

// Add default color scale conditional formatting
function AddDefaultColorScale(sheet) {
    const xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range.get("A5:C6"));
    sheet.Range.get("A5:C6").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("A5:C6").Style.Color = wasmModule.Color.get_Pink();
    const cf = xcfs.AddCondition();
    cf.FormatType = wasmModule.ConditionalFormatType.ColorScale;
    sheet.Range.get("A5").NumberValue = 4;
    sheet.Range.get("B5").NumberValue = 7;
    sheet.Range.get("C5").NumberValue = 10;
    sheet.Range.get("A6").NumberValue = 6;
    sheet.Range.get("B6").NumberValue = 9;
    sheet.Range.get("C6").NumberValue = 12;
}

// Add 3-color scale conditional formatting
function Add3ColorScale(sheet) {
    const xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range.get("A7:C8"));
    sheet.Range.get("A7:C8").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("A7:C8").Style.Color = wasmModule.Color.get_Green();
    const cf = xcfs.AddCondition();
    cf.FormatType = wasmModule.ConditionalFormatType.ColorScale;
    cf.ColorScale.MinValue.Type = wasmModule.ConditionValueType.Number;
    cf.ColorScale.MinValue.Value = 9;
    cf.ColorScale.MinColor = wasmModule.Color.get_Purple();
    sheet.Range.get("A7").NumberValue = 6;
    sheet.Range.get("B7").NumberValue = 9;
    sheet.Range.get("C7").NumberValue = 12;
    sheet.Range.get("A8").NumberValue = 8;
    sheet.Range.get("B8").NumberValue = 11;
    sheet.Range.get("C8").NumberValue = 14;
}

// Add data bar conditional formatting
function AddDataBar1(sheet) {
    const xcfs = sheet.ConditionalFormats.Add();
    xcfs.AddRange(sheet.Range.get("E1:G2"));
    sheet.Range.get("E1:G2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("E1:G2").Style.Color = wasmModule.Color.get_YellowGreen();
    const cf = xcfs.AddCondition();
    cf.FormatType = wasmModule.ConditionalFormatType.DataBar;
    cf.DataBar.BarColor = wasmModule.Color.get_Blue();
    cf.DataBar.MinPoint.Type = wasmModule.ConditionValueType.Percent;
    cf.DataBar.ShowValue = true;
    sheet.Range.get("E1").NumberValue = 4;
    sheet.Range.get("F1").NumberValue = 7;
    sheet.Range.get("G1").NumberValue = 10;
    sheet.Range.get("E2").NumberValue = 6;
    sheet.Range.get("F2").NumberValue = 9;
    sheet.Range.get("G2").NumberValue = 14;
}

// Add above average conditional formatting
function AddAboveAverage(sheet) {
    const conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range.get("A11:C12"));
    sheet.Range.get("A11:C12").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("A11:C12").Style.Color = wasmModule.Color.get_Tomato();
    const cf = conds.AddAverageCondition(wasmModule.AverageType.Above);
    cf.FillPattern = wasmModule.ExcelPatternType.Solid;
    cf.BackColor = wasmModule.Color.get_Pink();
    sheet.Range.get("A11").NumberValue = 10;
    sheet.Range.get("B11").NumberValue = 13;
    sheet.Range.get("C11").NumberValue = 16;
    sheet.Range.get("A12").NumberValue = 12;
    sheet.Range.get("B12").NumberValue = 15;
    sheet.Range.get("C12").NumberValue = 18;
}

// Add contains text conditional formatting
function AddContainsText(sheet) {
    const conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range.get("E5:G6"));
    sheet.Range.get("E5:G6").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("E5:G6").Style.Color = wasmModule.Color.get_LightBlue();
    const cf = conds.AddContainsTextCondition("abc");
    cf.FillPattern = wasmModule.ExcelPatternType.Solid;
    cf.BackColor = wasmModule.Color.get_Yellow();
    sheet.Range.get("E5").Text = "aa";
    sheet.Range.get("F5").Text = "abfd";
    sheet.Range.get("G5").Text = "aab";
    sheet.Range.get("E6").Text = "abc";
    sheet.Range.get("F6").Text = "cedf";
    sheet.Range.get("G6").Text = "abcd";
}

// Add time period conditional formatting for today
function AddTimePeriod_1(sheet) {
    let conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range.get("I1:K2"));
    sheet.Range.get("I1:K2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("I1:K2").Style.Color = wasmModule.Color.get_LightSlateGray();
    let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.Today);
    cf.FillPattern = wasmModule.ExcelPatternType.Solid;
    cf.BackColor = wasmModule.Color.get_Pink();
    let c = sheet.Range.get("I1");
    c.Value2 = wasmModule.DateTime.get_Now().AddDays(-8);
    c = sheet.Range.get("J1");
    c.Value2 = wasmModule.DateTime.get_Now().AddDays(-7);
    c = sheet.Range.get("K1");
    c.Value2 = wasmModule.DateTime.get_Now();
    c = sheet.Range.get("I2");
    c.Text = "Today";
    c = sheet.Range.get("J2");
    c.Value2 = wasmModule.DateTime.get_Now().AddDays(3);
    c = sheet.Range.get("K2");
    c.Value2 = wasmModule.DateTime.get_Now().AddMonths(2);
}

// Add duplicate values conditional formatting
function AddDuplicate(sheet) {
    let conds = sheet.ConditionalFormats.Add();
    conds.AddRange(sheet.Range.get("E23:G24"));
    sheet.Range.get("E23:G24").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
    sheet.Range.get("E23:G24").Style.Color = wasmModule.Color.get_LightSlateGray();
    let cf = conds.AddDuplicateValuesCondition();
    cf.FillPattern = wasmModule.ExcelPatternType.Solid;
    cf.BackColor = wasmModule.Color.get_Pink();
    let c = sheet.Range.get("E23");
    c.Text = "aa";
    c = sheet.Range.get("F23");
    c.Text = "bb";
    c = sheet.Range.get("G23");
    c.Text = "aa";
    c = sheet.Range.get("E24");
    c.Text = "bbb";
    c = sheet.Range.get("F24");
    c.Text = "bb";
    c = sheet.Range.get("G24");
    c.Text = "ccc";
}

// Main function to add all conditional formatting types
function AddConditionalFormattingForNewSheet(sheet) {
    AddDefaultIconSet(sheet);
    AddIconSet2(sheet);
    AddDefaultColorScale(sheet);
    Add3ColorScale(sheet);
    AddDataBar1(sheet);
    AddAboveAverage(sheet);
    AddContainsText(sheet);
    AddTimePeriod_1(sheet);
    AddDuplicate(sheet);
    sheet.AllocatedRange.ColumnWidth = 15;
    sheet.AllocatedRange.AutoFitRows();
}
```

---

# Excel Formula with Named Range
## Insert formula using named range in Excel
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();
const sheet = workbook.Worksheets.get(0);

//Set value
sheet.Range.get("A1").Value = "1";
sheet.Range.get("A2").Value = "1";

//Create a named range
const NamedRange = workbook.NameRanges.Add("NewNamedRange");
NamedRange.NameLocal = "=SUM(A1+A2)";

//Set the formula
sheet.Range.get("C1").Formula = "NewNamedRange";
```

---

# spire.xls javascript formula
## read formula from excel cell
```javascript
//Load the document
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: excelFileName});

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

//Get the formula value
const formula = sheet.Range.get("C14").Formula;
const value = sheet.Range.get("C14").FormulaNumberValue.toString();
```

---

# Excel Add-In Function Registration
## Register and use custom Add-In functions in Excel workbook
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Register AddIn function
workbook.AddInFunctions.Add(fileName, "TEST_UDF");
workbook.AddInFunctions.Add(fileName, "TEST_UDF1");

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

//Call AddIn function
sheet.Range.get("A1").Formula = "=TEST_UDF()";
sheet.Range.get("A2").Formula = "=TEST_UDF1()";
```

---

# spire.xls javascript formula
## remove formulas but keep values in excel cells
```javascript
//Loop through worksheets
for (let i = 0; i < workbook.Worksheets.Count; i++) {
  let sheet = workbook.Worksheets.get(i);
  //Loop through cells
  for (const cell of sheet.Range.Cells) {
    //If the cell contains formula, get the formula value, clear cell content, and then fill the formula value into the cell.
    if (cell.HasFormula) {
      const value = cell.FormulaValue;
      cell.Clear(wasmModule.ExcelClearOptions.ClearContent);
      cell.Value2 = wasmModule.String.Create(value);
    }
  }
}
```

---

# Excel SUBTOTAL Formulas in JavaScript
## Create and calculate SUBTOTAL formulas in an Excel spreadsheet using JavaScript
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

sheet.Range.get("A1").NumberValue = 1;
sheet.Range.get("A2").NumberValue = 2;
sheet.Range.get("A3").NumberValue = 3;
sheet.Range.get("B1").NumberValue = 4;
sheet.Range.get("B2").NumberValue = 5;
sheet.Range.get("B3").NumberValue = 6;
sheet.Range.get("C1").NumberValue = 7;
sheet.Range.get("C2").NumberValue = 8;
sheet.Range.get("C3").NumberValue = 9;

//Add SUBTOTAL formulas
sheet.Range.get("A5").Formula = "=SUBTOTAL(1,A1:C3)";
sheet.Range.get("B5").Formula = "=SUBTOTAL(2,A1:C3)";
sheet.Range.get("C5").Formula = "=SUBTOTAL(5,A1:C3)";

//Calculate Formulas
workbook.CalculateAllValue();
```

---

# spire.xls javascript array formulas
## use array formulas in excel
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get the first sheet
const sheet = workbook.Worksheets.get(0);

// Write array formula
sheet.Range.get("A5:C6").FormulaArray = "=LINEST(A1:A3,B1:C3,TRUE,TRUE)";

// Calculate Formulas
workbook.CalculateAllValue();
```

---

# Excel Array R1C1 Formula Implementation
## Demonstrates how to use array R1C1 formulas in Excel with JavaScript
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get the first sheet
const sheet = workbook.Worksheets.get(0);

// Add data to cells
sheet.Range.get("A1").NumberValue = 1;
sheet.Range.get("A2").NumberValue = 2;
sheet.Range.get("A3").NumberValue = 3;
sheet.Range.get("B1").NumberValue = 4;
sheet.Range.get("B2").NumberValue = 5;
sheet.Range.get("B3").NumberValue = 6;
sheet.Range.get("C1").NumberValue = 7;
sheet.Range.get("C2").NumberValue = 8;
sheet.Range.get("C3").NumberValue = 9;

// Add label for the formula result
sheet.Range.get("B4").Text = "Sum:";
sheet.Range.get("B4").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Right;

// Write array R1C1 formula
sheet.Range.get("C4").FormulaArrayR1C1 = "=SUM(R[-3]C[-2]:R[-1]C)";

// Calculate formulas
workbook.CalculateAllValue();
```

---

# spire.xls javascript r1c1 formula
## use array formula in excel
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get the first sheet
const sheet = workbook.Worksheets.get(0);

// Set cell values
sheet.Range.get("A1").NumberValue = 1;
sheet.Range.get("A2").NumberValue = 2;
sheet.Range.get("A3").NumberValue = 3;
sheet.Range.get("B1").NumberValue = 4;
sheet.Range.get("B2").NumberValue = 5;
sheet.Range.get("B3").NumberValue = 6;
sheet.Range.get("C1").NumberValue = 7;
sheet.Range.get("C2").NumberValue = 8;
sheet.Range.get("C3").NumberValue = 9;

// Write array formula
sheet.Range.get("A5:C6").FormulaArray = "=LINEST(A1:A3,B1:C3,TRUE,TRUE)";

// Calculate formulas
workbook.CalculateAllValue();
```

---

# spire.xls javascript formulas
## write various excel formulas to worksheet
```javascript
//Set column width 
sheet.SetColumnWidth(1, 32);
sheet.SetColumnWidth(2, 16);
sheet.SetColumnWidth(3, 16);		

//Set values
sheet.Range.get({row:1, column:1}).Value = "Examples of formulas :";
sheet.Range.get({row:3, column:1}).Value = "Test data:";
sheet.Range.get({row:3, column:2}).NumberValue = 7.3;
sheet.Range.get({row:3, column:3}).NumberValue = 5;
sheet.Range.get({row:3, column:4}).NumberValue = 8.2;
sheet.Range.get({row:3, column:5}).NumberValue = 4;
sheet.Range.get({row:3, column:6}).NumberValue = 3;
sheet.Range.get({row:3, column:7}).NumberValue = 11.3;

sheet.Range.get({row:4, column:1}).Value = "Formulas";
sheet.Range.get({row:4, column:2}).Value = "Results";

// String formula
sheet.Range.get({row:5, column:1}).NumberFormat="@";
sheet.Range.get({row:5, column:1}).Text = "=\"hello\"";
sheet.Range.get({row:5, column:2}).Formula = "=\"hello\"";
sheet.Range.get({row:5, column:3}).Formula = "=\"" + '\u4f60\u597d' + "\"";

// Numeric formulas
sheet.Range.get({row:6, column:1}).NumberFormat="@";
sheet.Range.get({row:6, column:1}).Text = "=300";
sheet.Range.get({row:6, column:2}).Formula = "=300";

sheet.Range.get({row:7, column:1}).NumberFormat="@";
sheet.Range.get({row:7, column:1}).Text = "=3389.639421";
sheet.Range.get({row:7, column:2}).Formula = "=3389.639421";

// Boolean formula
sheet.Range.get({row:8, column:1}).NumberFormat="@";
sheet.Range.get({row:8, column:1}).Text = "=false";
sheet.Range.get({row:8, column:2}).Formula = "=false";

// Arithmetic formulas
sheet.Range.get({row:9, column:1}).NumberFormat="@";
sheet.Range.get({row:9, column:1}).Text = "=1+2+3+4+5-6-7+8-9";
sheet.Range.get({row:9, column:2}).Formula = "=1+2+3+4+5-6-7+8-9";

sheet.Range.get({row:10, column:1}).NumberFormat="@";
sheet.Range.get({row:10, column:1}).Text = "=33*3/4-2+10";
sheet.Range.get({row:10, column:2}).Formula = "=33*3/4-2+10";

// Cell reference formula
sheet.Range.get({row:11, column:1}).NumberFormat="@";
sheet.Range.get({row:11, column:1}).Text = "=Sheet1!$B$3";
sheet.Range.get({row:11, column:2}).Formula = "=Sheet1!$B$3";

// Function formulas
sheet.Range.get({row:12, column:1}).NumberFormat="@";
sheet.Range.get({row:12, column:1}).Text = "=AVERAGE(Sheet1!$D$3:G$3)";
sheet.Range.get({row:12, column:2}).Formula = "=AVERAGE(Sheet1!$D$3:G$3)";

sheet.Range.get({row:13, column:1}).NumberFormat="@";
sheet.Range.get({row:13, column:1}).Text = "=Count(3,5,8,10,2,34)";
sheet.Range.get({row:13, column:2}).Formula = "=Count(3,5,8,10,2,34)";

// Date and time formulas
sheet.Range.get({row:14, column:1}).NumberFormat="@";
sheet.Range.get({row:14, column:1}).Text = "=NOW()";
sheet.Range.get({row:14, column:2}).Formula = "=NOW()";
sheet.Range.get({row:14, column:2}).Style.NumberFormat = "yyyy-MM-DD";

sheet.Range.get({row:15, column:1}).NumberFormat="@";
sheet.Range.get({row:15, column:1}).Text = "=SECOND(11)";
sheet.Range.get({row:15, column:2}).Formula = "=SECOND(11)";

sheet.Range.get({row:16, column:1}).NumberFormat="@";
sheet.Range.get({row:16, column:1}).Text = "=MINUTE(12)";
sheet.Range.get({row:16, column:2}).Formula = "=MINUTE(12)";

sheet.Range.get({row:17, column:1}).NumberFormat="@";
sheet.Range.get({row:17, column:1}).Text = "=MONTH(9)";
sheet.Range.get({row:17, column:2}).Formula = "=MONTH(9)";

sheet.Range.get({row:18, column:1}).NumberFormat="@";
sheet.Range.get({row:18, column:1}).Text = "=DAY(10)";
sheet.Range.get({row:18, column:2}).Formula = "=DAY(10)";

sheet.Range.get({row:19, column:1}).NumberFormat="@";
sheet.Range.get({row:19, column:1}).Text = "=TIME(4,5,7)";
sheet.Range.get({row:19, column:2}).Formula = "=TIME(4,5,7)";

sheet.Range.get({row:20, column:1}).NumberFormat="@";
sheet.Range.get({row:20, column:1}).Text = "=DATE(6,4,2)";
sheet.Range.get({row:20, column:2}).Formula = "=DATE(6,4,2)";

// Math formulas
sheet.Range.get({row:21, column:1}).NumberFormat="@";
sheet.Range.get({row:21, column:1}).Text = "=RAND()";
sheet.Range.get({row:21, column:2}).Formula = "=RAND()";

sheet.Range.get({row:22, column:1}).NumberFormat="@";
sheet.Range.get({row:22, column:1}).Text = "=HOUR(12)";
sheet.Range.get({row:22, column:2}).Formula = "=HOUR(12)";

sheet.Range.get({row:23, column:1}).NumberFormat="@";
sheet.Range.get({row:23, column:1}).Text = "=MOD(5,3)";
sheet.Range.get({row:23, column:2}).Formula = "=MOD(5,3)";

sheet.Range.get({row:24, column:1}).NumberFormat="@";
sheet.Range.get({row:24, column:1}).Text = "=WEEKDAY(3)";
sheet.Range.get({row:24, column:2}).Formula = "=WEEKDAY(3)";

sheet.Range.get({row:25, column:1}).NumberFormat="@";
sheet.Range.get({row:25, column:1}).Text = "=YEAR(23)";
sheet.Range.get({row:25, column:2}).Formula = "=YEAR(23)";

// Logical formulas
sheet.Range.get({row:26, column:1}).NumberFormat="@";
sheet.Range.get({row:26, column:1}).Text = "=NOT(true)";
sheet.Range.get({row:26, column:2}).Formula = "=NOT(true)";

sheet.Range.get({row:27, column:1}).NumberFormat="@";
sheet.Range.get({row:27, column:1}).Text = "=OR(true)";
sheet.Range.get({row:27, column:2}).Formula = "=OR(true)";

sheet.Range.get({row:28, column:1}).NumberFormat="@";
sheet.Range.get({row:28, column:1}).Text = "=AND(TRUE)";
sheet.Range.get({row:28, column:2}).Formula = "=AND(TRUE)";

// Text formulas
sheet.Range.get({row:29, column:1}).NumberFormat="@";
sheet.Range.get({row:29, column:1}).Text = "=VALUE(30)";
sheet.Range.get({row:29, column:2}).Formula = "=VALUE(30)";

sheet.Range.get({row:30, column:1}).NumberFormat="@";
sheet.Range.get({row:30, column:1}).Text = "=LEN(\"world\")";
sheet.Range.get({row:30, column:2}).Formula = "=LEN(\"world\")";

sheet.Range.get({row:31, column:1}).NumberFormat="@";
sheet.Range.get({row:31, column:1}).Text = "=MID(\"world\",4,2)";
sheet.Range.get({row:31, column:2}).Formula = "=MID(\"world\",4,2)";

// Advanced math formulas
sheet.Range.get({row:32, column:1}).NumberFormat="@";
sheet.Range.get({row:32, column:1}).Text = "=ROUND(7,3)";
sheet.Range.get({row:32, column:2}).Formula = "=ROUND(7,3)";

sheet.Range.get({row:33, column:1}).NumberFormat="@";
sheet.Range.get({row:33, column:1}).Text = "=SIGN(4)";
sheet.Range.get({row:33, column:2}).Formula = "=SIGN(4)";

sheet.Range.get({row:34, column:1}).NumberFormat="@";
sheet.Range.get({row:34, column:1}).Text = "=INT(200)";
sheet.Range.get({row:34, column:2}).Formula = "=INT(200)";

sheet.Range.get({row:35, column:1}).NumberFormat="@";
sheet.Range.get({row:35, column:1}).Text = "=ABS(-1.21)";
sheet.Range.get({row:35, column:2}).Formula = "=ABS(-1.21)";

sheet.Range.get({row:36, column:1}).NumberFormat="@";
sheet.Range.get({row:36, column:1}).Text = "=LN(15)";
sheet.Range.get({row:36, column:2}).Formula = "=LN(15)";

sheet.Range.get({row:37, column:1}).NumberFormat="@";
sheet.Range.get({row:37, column:1}).Text = "=EXP(20)";
sheet.Range.get({row:37, column:2}).Formula = "=EXP(20)";

sheet.Range.get({row:38, column:1}).NumberFormat="@";
sheet.Range.get({row:38, column:1}).Text = "=SQRT(40)";
sheet.Range.get({row:38, column:2}).Formula = "=SQRT(40)";

sheet.Range.get({row:39, column:1}).NumberFormat="@";
sheet.Range.get({row:39, column:1}).Text = "=PI()";
sheet.Range.get({row:39, column:2}).Formula = "=PI()";

sheet.Range.get({row:40, column:1}).NumberFormat="@";
sheet.Range.get({row:40, column:1}).Text = "=COS(9)";
sheet.Range.get({row:40, column:2}).Formula = "=COS(9)";

sheet.Range.get({row:41, column:1}).NumberFormat="@";
sheet.Range.get({row:41, column:1}).Text = "=SIN(45)";
sheet.Range.get({row:41, column:2}).Formula = "=SIN(45)";

// Statistical formulas
sheet.Range.get({row:42, column:1}).NumberFormat="@";
sheet.Range.get({row:42, column:1}).Text = "=MAX(10,30)";
sheet.Range.get({row:42, column:2}).Formula = "=MAX(10,30)";

sheet.Range.get({row:43, column:1}).NumberFormat="@";
sheet.Range.get({row:43, column:1}).Text = "=MIN(5,7)";
sheet.Range.get({row:43, column:2}).Formula = "=MIN(5,7)";

sheet.Range.get({row:44, column:1}).NumberFormat="@";
sheet.Range.get({row:44, column:1}).Text = "=AVERAGE(12,45)";
sheet.Range.get({row:44, column:2}).Formula = "=AVERAGE(12,45)";

sheet.Range.get({row:45, column:1}).NumberFormat="@";
sheet.Range.get({row:45, column:1}).Text = "=SUM(18,29)";
sheet.Range.get({row:45, column:2}).Formula = "=SUM(18,29)";

// Conditional formula
sheet.Range.get({row:46, column:1}).NumberFormat="@";
sheet.Range.get({row:46, column:1}).Text = "=IF(4,2,2)";
sheet.Range.get({row:46, column:2}).Formula = "=IF(4,2,2)";

// Advanced formula
sheet.Range.get({row:47, column:1}).NumberFormat="@";
sheet.Range.get({row:47, column:1}).Text = "=SUBTOTAL(3,Sheet1!B2:E3)";
sheet.Range.get({row:47, column:2}).Formula = "=SUBTOTAL(3,Sheet1!B2:E3)";
```

---

# Excel Header Footer Font Modification
## Change font and size for header and footer in Excel file
```javascript
// Get the first worksheet
const sheet = workbook.Worksheets.get(0);

// Set the new font and size for the header and footer
let text = sheet.PageSetup.LeftHeader;
text = "&\"Arial Unicode MS\"&18 Header Footer Sample by Spire.XLS ";
sheet.PageSetup.LeftHeader = text;
sheet.PageSetup.RightFooter = text;
```

---

# Excel Header and Footer Configuration
## Set different headers and footers for odd and even pages
```javascript
// Set the different header footer for Odd and Even pages
sheet.PageSetup.DifferentOddEven = 1;

// Set the header with font, size, bold, and color
sheet.PageSetup.OddHeaderString = "&\"Arial\"&12&B&KFFC000 Odd_Header";
sheet.PageSetup.OddFooterString = "&\"Arial\"&12&B&KFFC000 Odd_Footer";
sheet.PageSetup.EvenHeaderString = "&\"Arial\"&12&B&KFF0000 Even_Header";
sheet.PageSetup.EvenFooterString = "&\"Arial\"&12&B&KFF0000 Even_Footer";

sheet.ViewMode = wasmModule.ViewMode.Layout;
```

---

# Excel Header Footer Configuration
## Set different header and footer for the first page in Excel
```javascript
// Set the value to show the headers/footers for first page are different from the other pages
sheet.PageSetup.DifferentFirst = 1;

// Set the header and footer for the first page
sheet.PageSetup.FirstHeaderString = "Different First page";
sheet.PageSetup.FirstFooterString = "Different First footer";

// Set the other pages' header and footer
sheet.PageSetup.LeftHeader = "Demo of Spire.XLS";
sheet.PageSetup.CenterFooter = "Footer by Spire.XLS";
```

---

# Excel Header and Footer with Images
## Add images to header and footer in Excel worksheets
```javascript
// Load an image 
const image = wasmModule.Stream.CreateByFile( imgFileName);

// Set the image header
sheet.PageSetup.LeftHeaderImage = image;
sheet.PageSetup.LeftHeader = "&G";

// Set the image footer
sheet.PageSetup.CenterFooterImage = image;
sheet.PageSetup.CenterFooter = "&G";

// Set the view mode of the sheet
sheet.ViewMode = wasmModule.ViewMode.Layout;
```

---

# spire.xls javascript header footer
## set header and footer in Excel worksheet
```javascript
// Set left header,"Arial Unicode MS" is font name, "18" is font size.
worksheet.PageSetup.LeftHeader = "&\"Arial Unicode MS\"&14 Spire.XLS for .NET ";

// Set center footer
worksheet.PageSetup.CenterFooter = "Footer Text";

worksheet.ViewMode = wasmModule.ViewMode.Layout;
```

---

# spire.xls javascript hyperlink
## add hyperlink to text in excel workbook
```javascript
//Get the first sheet
const sheet = workbook.Worksheets.get(0);

//Add url link
const UrlLink = sheet.HyperLinks.Add({range:sheet.Range.get("D10")});
UrlLink.TextToDisplay = sheet.Range.get("D10").Text;
UrlLink.Type = wasmModule.HyperLinkType.Url;
UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago";

//Add email link
const MailLink = sheet.HyperLinks.Add({range:sheet.Range.get("E10")});
MailLink.TextToDisplay = sheet.Range.get("E10").Text;
MailLink.Type = wasmModule.HyperLinkType.Url;
MailLink.Address = "mailto:Amor.Aqua@gmail.com";
```

---

# Spire.XLS JavaScript Image Hyperlink
## Add a hyperlink to an image in an Excel worksheet
```javascript
//Insert an image to a specific cell
let imgFileName = 'logo.png';
let picture = sheet.Pictures.Add({topRow:2, leftColumn:1, fileName:imgFileName});
//Add a hyperlink to the image
picture.SetHyperLink("https://www.e-iceblue.com/Misc/about-us.html", true);
```

---

# Get Hyperlink Types in Excel
## Extract hyperlink addresses and types from an Excel worksheet
```javascript
//Get the first worksheet
const sheet = workbook.Worksheets.get(0);

//Iterate all hyperlinks
const sb = [];
const hyperlinks = sheet.HyperLinks;
for(let i=0; i<hyperlinks.Count; i++) {
    const item = hyperlinks.get(i);
    //Get hyperlink address
    const address = item.Address;
    //Get hyperlink type
    const type = item.Type;
    sb.push(`Link address: ${address}`);
    sb.push(`Link type: ${type}`);
    sb.push("");
}
```

---

# Spire.XLS JavaScript Hyperlink
## Add hyperlink to external file in Excel
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

const range = sheet.Range.get("A1");

//Add hyperlink in the range
const hyperlink = sheet.HyperLinks.Add({range:range});

//Set the link type
hyperlink.Type = wasmModule.HyperLinkType.File;

//Set the display text
hyperlink.TextToDisplay = "Link To External File";

//Set file address
hyperlink.Address ="https://www.baidu.com/img/PCtm_d9c8750bed0b3c7d089fa7d55720d6cf.png";

//Dispose
workbook.Dispose();
```

---

# spire.xls javascript hyperlinks
## add hyperlink to other sheet cell
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

const range = sheet.Range.get("A1");

//Add hyperlink in the range
const hyperlink = sheet.HyperLinks.Add({range:range});

//Set the link type
hyperlink.Type = wasmModule.HyperLinkType.Workbook;

//Set the display text
hyperlink.TextToDisplay = "Link to Sheet2 cell C5";

//Set the address
hyperlink.Address = "Sheet2!C5";
```

---

# spire.xls javascript hyperlink
## modify hyperlink in excel worksheet
```javascript
//Get the collection of all hyperlinks in the worksheet
const sheet = workbook.Worksheets.get(0);

//Change the values of TextToDisplay and Address property
const links = sheet.HyperLinks;
links.get(0).TextToDisplay = "Spire.XLS for .NET";
links.get(0).Address = "https://www.e-iceblue.com/";
```

---

# Excel Hyperlink Reader
## Extract hyperlinks from Excel cells
```javascript
// Get the first sheet
const sheet = workbook.Worksheets.get(0);
// Get hyperlink addresses
const address1 = sheet.HyperLinks.get(0).Address;
const address2 = sheet.HyperLinks.get(1).Address;
```

---

# Excel Hyperlink Removal
## Remove hyperlinks from Excel worksheet while preserving text
```javascript
//Get the first worksheet
const sheet = workbook.Worksheets.get(0);

//Get the collection of all hyperlinks in the worksheet
const links = sheet.HyperLinks;

//Remove all link content
sheet.Range.get("B1").ClearAll();
sheet.Range.get("B2").ClearAll();
sheet.Range.get("B3").ClearAll();

//Remove hyperlink and keep link text
sheet.HyperLinks.RemoveAt(0);
sheet.HyperLinks.RemoveAt(0);
sheet.HyperLinks.RemoveAt(0);
```

---

# Retrieve External File Hyperlinks in Excel
## Extract and display external file hyperlinks from Excel worksheets
```javascript
//Load the document
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: excelFileName});

//Get the first sheet
const sheet = workbook.Worksheets.get(0);

const content = [];

//Retrieve external file hyperlinks.
const hyperlinks = sheet.HyperLinks;
for(let i=0; i<hyperlinks.Count; i++) {
	const item = hyperlinks.get(i);
	const address = item.Address;
	const sheetName = item.Range.WorksheetName;
	const range = item.Range;
	content.push(`Cell[${range.Row},${range.Column}] in sheet "${sheetName}" contains File URL: ${address}`);
}
```

---

# spire.xls javascript hyperlinks
## write hyperlinks to excel cells
```javascript
//Get the first sheet
const sheet = workbook.Worksheets.get(0);

//Set links
sheet.Range.get("B9").Text = "Home page";
const hylink1 = sheet.HyperLinks.Add({range:sheet.Range.get("B10")});
hylink1.Type = wasmModule.HyperLinkType.Url;
hylink1.Address = "http://www.e-iceblue.com";

sheet.Range.get("B11").Text = "Support";
const hylink2 = sheet.HyperLinks.Add({range:sheet.Range.get("B12")});
hylink2.Type = wasmModule.HyperLinkType.Url;
hylink2.Address = "mailto:support@e-iceblue.com";

sheet.Range.get("B13").Text = "Forum";
const hylink3 = sheet.HyperLinks.Add({range:sheet.Range.get("B14")});
hylink3.Type = wasmModule.HyperLinkType.Url;
hylink3.Address = "https://www.e-iceblue.com/forum/";
```

---

# Excel Variable Array with Marker Designer
## Add variable array to Excel workbook using marker designer functionality
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get the first worksheet
const sheet = workbook.Worksheets.get(0);

// Set marker designer field in cell A1
sheet.Range.get("A1").Value = "&=Array";

// Fill Array
workbook.MarkerDesigner.AddArray("Array",
  [wasmModule.String.Create("Spire.Xls"),
  wasmModule.String.Create("Spire.Doc"),
  wasmModule.String.Create("Spire.PDF"),
  wasmModule.String.Create("Spire.Presentation"),
  wasmModule.String.Create("Spire.Email")]);
workbook.MarkerDesigner.Apply();
workbook.CalculateAllValue();

// AutoFit
sheet.AllocatedRange.AutoFitRows();
sheet.AllocatedRange.AutoFitColumns();
```

---

# Excel Named Range Formatting
## Format cells in a named range by setting color and bold font
```javascript
// Get specific named range by index
let NamedRange = book.NameRanges.get(0);

// Get the cell range of the named range
let range = NamedRange.RefersToRange;

// Set color for the range
range.Style.Color = wasmModule.Color.get_Yellow();

// Set the font as bold
range.Style.Font.IsBold = true;
```

---

# Get All Named Ranges in Excel
## This example demonstrates how to retrieve all named ranges from an Excel workbook
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

let sb = [];
// Get all named ranges
let ranges = book.NameRanges;
for (let i = 0; i < ranges.Count; i++) {
  let nameRange = ranges.get(i);
  sb.push(nameRange.Name);
}

// Define the output file name
const outputFileName = "GetAllNamedRange.txt";
// Save result file
wasmModule.FS.writeFile(outputFileName, sb.join("\n"));
```

---

# Getting Named Range Address
## Retrieve the address of a named range in an Excel workbook
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Get specific named range by index
let NamedRange = book.NameRanges.get(0);

// Get the address of the named range
let address = NamedRange.RefersToRange.RangeAddress;

// Store the result
let sb = [];
sb.push(
  `The address of the named range ${NamedRange.Name} is ${address}`
);
```

---

# Spire.XLS JavaScript Named Ranges
## Get specific named range by index or name
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Get specific named range by index
let name1 = book.NameRanges.get(1).Name;

// Get specific named range by name
let name2 = book.NameRanges.get({ name: "NameRange3" }).Name;

// Clean up resources
book.Dispose();
```

---

# Merge Named Range Cells
## Demonstrates how to merge cells within a named range in an Excel workbook
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Get specific named range by index
let namedRange = book.NameRanges.get(0);

// Get the range of the named range
let range = namedRange.RefersToRange;

// Merge cells
range.Merge();

// Clean up resources
book.Dispose();
```

---

# spire.xls javascript named ranges
## create named range in excel
```javascript
// Creating a named range
let namedRange = book.NameRanges.Add("NewNamedRange");

// Setting the range of the named range
namedRange.RefersToRange = sheet.Range.get("A8:E12");
```

---

# Spire.XLS JavaScript Named Range Management
## Remove named ranges from Excel workbook
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Remove the named range by index
book.NameRanges.RemoveAt(0);

// Remove the named range by name
book.NameRanges.Remove("NameRange2");

// Clean up resources
book.Dispose();
```

---

# Excel Named Range Renaming
## Rename a named range in an Excel workbook
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Rename the named range
book.NameRanges.get(0).Name = "RenameRange";

// Save the workbook to the specified path
book.SaveToFile({
  fileName: outputFileName,
  version: wasmModule.ExcelVersion.Version2010,
});
```

---

# Excel Named Range Management
## Create a scoped named range in Excel worksheet
```javascript
// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Add range name
let namedRange = sheet.Names.Add("Range1");

// Define the range
namedRange.RefersToRange = sheet.Range.get("A1:D19");
```

---

# spire.xls javascript named range
## set formula with named range
```javascript
// Create a named range
let namedRange = book.NameRanges.Add("MyNamedRange");

// Refers to range
namedRange.RefersToRange = sheet.Range.get("B10:B12");

// Set the formula of range to named range
sheet.Range.get("B13").Formula = "=SUM(MyNamedRange)";

// Set value of ranges
sheet.Range.get("B10").Value2 = wasmModule.Int32.Create(10);
sheet.Range.get("B11").Value2 = wasmModule.Int32.Create(20);
sheet.Range.get("B12").Value2 = wasmModule.Int32.Create(30);
```

---

# spire.xls javascript ole objects
## extract OLE objects from Excel worksheet
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Extract ole objects
if (sheet.HasOleObjects) {
  for (let obj of sheet.OleObjects) {
    let type = obj.ObjectType;
    // Word document
    if (type === wasmModule.OleObjectType.WordDocument) {
      // Define the output file name
      const outputFileName = "ExtractOLEObjects.docx";
      // Process the Word document OLE data
      topostMessageDoc(outputFileName, obj.OleData);
    }
    // Pdf document
    if (type === wasmModule.OleObjectType.AdobeAcrobatDocument) {
      // Define the output file name
      const outputFileName = "ExtractOLEObjects.pdf";
      // Process the PDF document OLE data
      topostMessagePdf(outputFileName, obj.OleData);
    }
    // Ppt document
    if (type === wasmModule.OleObjectType.PowerPointSlide) {
      // Define the output file name
      const outputFileName = "ExtractOLEObjects.pptx";
      // Process the PowerPoint document OLE data
      topostMessagePpt(outputFileName, obj.OleData);
    }
  }
}

// Clean up resources
book.Dispose();
```

---

# Insert OLE Objects in Excel
## This code demonstrates how to insert OLE objects into an Excel worksheet
```javascript
// Insert OLE object
let book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);
let worksheet = book.Worksheets.get(0);
worksheet.PageSetup.LeftMargin = 0;
worksheet.PageSetup.RightMargin = 0;
worksheet.PageSetup.TopMargin = 0;
worksheet.PageSetup.BottomMargin = 0;
// Convert worksheet to image
let image = worksheet.ToImage(1, 1, 19, 5);
// Clean up resources
book.Dispose();
// Add OLE object
let oleObject = ws.OleObjects.Add(
  excelFileName,
  image,
  wasmModule.OleLinkType.Embed
);
oleObject.Location = ws.Range.get("B4");
oleObject.ObjectType = wasmModule.OleObjectType.ExcelWorksheet;
```

---

# Insert WAV File as OLE Object in Excel
## This code demonstrates how to insert a WAV file as an OLE object into an Excel worksheet
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Add OLE object
let fs = wasmModule.Stream.CreateByFile(pngFileName);
let oleObject = sheet.OleObjects.Add(
  wavFileName,
  fs,
  wasmModule.OleLinkType.Embed
);

// Set the object location
oleObject.Location = sheet.Range.get("B4");
// Set the object type as package
oleObject.ObjectType = wasmModule.OleObjectType.Package;
```

---

# Excel Paper Dimensions Retrieval
## Get dimensions of different paper sizes in Excel

```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
// Get the first worksheet.
let sheet = book.Worksheets.get(0);

// Get the dimensions of A2 paper.
sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.A2Paper;
let a2Width = sheet.PageSetup.PageWidth;
let a2Height = sheet.PageSetup.PageHeight;

// Get the dimensions of A3 paper.
sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperA3;
let a3Width = sheet.PageSetup.PageWidth;
let a3Height = sheet.PageSetup.PageHeight;

// Get the dimensions of A4 paper.
sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperA4;
let a4Width = sheet.PageSetup.PageWidth;
let a4Height = sheet.PageSetup.PageHeight;

// Get the dimensions of paper letter.
sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperLetter;
let letterWidth = sheet.PageSetup.PageWidth;
let letterHeight = sheet.PageSetup.PageHeight;
```

---

# Spire.XLS JavaScript Get Excel Version
## Retrieve Excel file version information
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Get the version
let version = book.Version;
```

---

# Excel Page Setup Order Type
## Set page order type in Excel worksheet
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});

// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Get the reference of the PageSetup of the worksheet
let pageSetup = sheet.PageSetup;

// Set the order type of the pages to over then down
pageSetup.Order = wasmModule.OrderType.OverThenDown;
```

---

# Excel Page Setup in JavaScript
## Set paper size and page order in Excel files
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();

// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Get the reference of the PageSetup of the worksheet
let pageSetup = sheet.PageSetup;

// Set the order type of the pages to over then down
pageSetup.Order = wasmModule.OrderType.OverThenDown;
```

---

# Excel Page Setup First Page Number
## Set the first page number for Excel worksheet
```javascript
// Get the first worksheet.
let sheet = book.Worksheets.get(0);

// Set the first page number of the worksheet pages.
sheet.PageSetup.FirstPageNumber = 2;
```

---

# Excel Page Setup Header and Footer Margins
## Set header and footer margins in Excel worksheet
```javascript
// Get the PageSetup object of the first worksheet.
let pageSetup = sheet.PageSetup;

// Set the margins of header and footer.
pageSetup.HeaderMarginInch = 2;
pageSetup.FooterMarginInch = 2;
```

---

# Excel Page Margin Setup
## Set page margins for Excel worksheets
```javascript
// Get the PageSetup object of the first worksheet
let pageSetup = sheet.PageSetup;

// Set bottom, left, right, and top page margins
pageSetup.BottomMargin = 2;
pageSetup.LeftMargin = 1;
pageSetup.RightMargin = 1;
pageSetup.TopMargin = 3;
```

---

# Excel Printing Options Setup
## Configure various printing options for Excel worksheets
```javascript
// Get the first worksheet.
let sheet = book.Worksheets.get(0);

// Get the reference of the PageSetup of the worksheet.
let pageSetup = sheet.PageSetup;

// Allow to print gridlines.
pageSetup.IsPrintGridlines = true;

// Allow to print row/column headings.
pageSetup.IsPrintHeadings = true;

// Allow to print worksheet in black & white mode.
pageSetup.BlackAndWhite = true;

// Allow to print comments as displayed on worksheet.
pageSetup.PrintComments = wasmModule.PrintCommentType.InPlace;

// Allow to print worksheet with draft quality.
pageSetup.Draft = true;

// Allow to print cell errors as N/A.
pageSetup.PrintErrors = wasmModule.PrintErrorsType.NA;
```

---

# Excel Page Setup
## Set page orientation in Excel worksheet
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
  fileName: excelFileName,
  version: wasmModule.ExcelVersion.Version2010,
});
// Get the first worksheet.
let sheet = book.Worksheets.get(0);

// Set the page orientation to Landscape.
sheet.PageSetup.Orientation = wasmModule.PageOrientationType.Landscape;
```

---

# Excel Print Area Setup
## Set print area for Excel worksheet
```javascript
//Get the reference of the PageSetup of the worksheet.
let pageSetup = sheet.PageSetup;

//Specify the cells range of the print area.
pageSetup.PrintArea = "A1:E5";
```

---

# Excel Print Quality Setup
## Set the print quality of an Excel worksheet
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
    fileName: excelFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Set the print quality of the worksheet to 180 dpi
sheet.PageSetup.PrintQuality = 180;
```

---

# Excel Print Title Setup
## Set print title columns and rows in Excel worksheet
```javascript
// Get the PageSetup of the worksheet
let pageSetup = sheet.PageSetup;

//Define column numbers A & B as title columns.
pageSetup.PrintTitleColumns = '$A:$B';

//Defining row numbers 1 & 2 as title rows.
pageSetup.PrintTitleRows = '$1:$2';
```

---

# spire.xls javascript page setup
## Set worksheet FitToPage property
```javascript
//Get the first worksheet.
let sheet = book.Worksheets.get(0);

let pageSetup = sheet.PageSetup;

//Set the FitToPagesTall property.
pageSetup.FitToPagesTall = 1;

//Set the FitToPagesWide property.
pageSetup.FitToPagesWide = 1;
```

---

# Excel Page Setup Centering
## Set worksheet to center horizontally and vertically on page when printed
```javascript
// Get the PageSetup object of the first page
let pageSetup = sheet.PageSetup;

// Set the worksheet center on page
pageSetup.CenterHorizontally = true;
pageSetup.CenterVertically = true;
```

---

# Pivot Table Data Source Modification
## Change the data source of a pivot table in an Excel workbook
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
    fileName: excelFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Define the range of cells to be used as the new data source
let Range = sheet.Range.get('A1:C15');

// Get the first pivot table from the second worksheet
let table = book.Worksheets.get(1).PivotTables.get(0);

// Change data source
table.ChangeDataSource(Range);
table.Cache.IsRefreshOnLoad = false;
```

---

# Clear Pivot Fields in Excel
## This code demonstrates how to clear all data fields in a pivot table
```javascript
// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
// Get the first pivot table from the sheet
let pt = sheet.PivotTables.get(0);

// Clear all the data fields
pt.DataFields.Clear();
// Calculate the pivot table data
pt.CalculateData();
```

---

# Pivot Table Consolidation Functions
## Applying Average and Max consolidation functions to pivot table data fields
```javascript
// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
let pt = sheet.PivotTables.get(0);
// Apply Average consolidation function to first data field
pt.DataFields.get(0).Subtotal = wasmModule.SubtotalTypes.Average;
// Apply Max consolidation function to second data field
pt.DataFields.get(1).Subtotal = wasmModule.SubtotalTypes.Max;
// Calculate the pivot table data
pt.CalculateData();
```

---

# Spire.XLS JavaScript PivotTable
## Create a PivotTable in Excel
```javascript
//Add a PivotTable to the worksheet
let dataRange = sheet.Range.get('A1:C7');
let cache = book.PivotCaches.Add({ range: dataRange });
let pt = sheet.PivotTables.Add('Pivot Table', sheet.Range.get({ row: 10, column: 5 }), cache);

//Drag the fields to the row area.
let pf = pt.PivotFields.get_Item('Product');
pf.Axis = wasmModule.AxisTypes.Row;
let pf2 = pt.PivotFields.get_Item('Month');
pf2.Axis = wasmModule.AxisTypes.Row;
//Drag the field to the data area.
pt.DataFields.Add(pt.PivotFields.get_Item('Count'), 'SUM of Count', wasmModule.SubtotalTypes.Sum);
//Set PivotTable style
pt.BuiltInStyle = wasmModule.PivotBuiltInStyles.PivotStyleMedium12;

//Autofit columns generated by the pivotTable
pt.CalculateData();
sheet.AutoFitColumn(5);
sheet.AutoFitColumn(6);
```

---

# Disable Pivot Table Ribbon
## This code demonstrates how to disable the pivot table ribbon in an Excel file
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
    fileName: excelFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });

let pt = sheet.PivotTables.get(0);
// Disable ribbon for this pivot table
pt.EnableWizard = false;
```

---

# spire.xls javascript pivot table
## expand or collapse rows in pivot table
```javascript
//Get the first worksheet.
let sheet = book.Worksheets.get(0);

//Get the data in Pivot Table.
let pivotTable = sheet.PivotTables.get(0);

//Calculate Data.
pivotTable.CalculateData();

//Collapse the rows.
pivotTable.PivotFields.get_Item('Vendor No').HideItemDetail({
    itemValue: '3501',
    isHiddenDetail: true,
});

//Expand the rows.
pivotTable.PivotFields.get_Item('Vendor No').HideItemDetail({
    itemValue: '3502',
    isHiddenDetail: false,
});
```

---

# spire.xls javascript pivot table
## format data field in pivot table
```javascript
// Access the PivotTable
let pt = sheet.PivotTables.get(0);
// Access the data field
let pivotDataField = pt.DataFields.get(0);
// Set data display format
pivotDataField.ShowDataAs = wasmModule.PivotFieldFormatType.PercentageOfColumn;
```

---

# spire.xls javascript pivot table
## format pivot table appearance
```javascript
// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
// Get the first pivot table from the worksheet
let pt = sheet.PivotTables.get(0);

// Format appearance
pt.BuiltInStyle = wasmModule.PivotBuiltInStyles.PivotStyleLight10;
// Enable the display of grid drop zone in the pivot table
pt.Options.ShowGridDropZone = true;
// Set the row layout type to Tabular in the pivot table
pt.Options.RowLayout = wasmModule.PivotTableLayoutType.Tabular;
```

---

# Get PivotTable Refresh Information
## Extract refresh date and user information from a PivotTable in an Excel worksheet
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
    fileName: excelFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Get first worksheet of the workbook
let worksheet = book.Worksheets.get(0);

// Get the first pivot table
let pivotTable = worksheet.PivotTables.get(0);

// Get the refreshed information
let dateTime = pivotTable.Cache.RefreshDate;
let refreshedBy = pivotTable.Cache.RefreshedBy;

// Set string format for displaying
let result = 'Pivot table refreshed by:  ' + refreshedBy + '\r\nPivot table refreshed date: ' + dateTime.ToString();
```

---

# Excel Pivot Table Refresh
## Refresh pivot table data in Excel using JavaScript
```javascript
// Get the second worksheet
let sheet = book.Worksheets.get(1);

// Update the data source of PivotTable
sheet.Range.get('D2').Value = '999';

// Get the PivotTable that was built on the data source
let pt = book.Worksheets.get(0).PivotTables.get(0);

// Refresh the data of PivotTable
pt.Cache.IsRefreshOnLoad = true;
```

---

# Pivot Table Item Labels
## Repeat item labels in pivot table
```javascript
// Create a pivot cache using the data range
let cache = book.PivotCaches.Add({ range: dataRange });
// Add a pivot table to the pivot sheet using the pivot cache
let pt = sheet2.PivotTables.Add('Pivot Table', sheet.Range.get('A1'), cache);

// Set the VendorNo field as a row field and specify its header caption
let r1 = pt.PivotFields.get_Item('VendorNo');
r1.Axis = wasmModule.AxisTypes.Row;
pt.Options.RowHeaderCaption = 'VendorNo';
r1.Subtotals = wasmModule.SubtotalTypes.None;

// Enable repeating item labels for the VendorNo field
r1.RepeatItemLabels = true;

// Enable repeating item labels for the OnHand field
pt.PivotFields.get_Item('OnHand').RepeatItemLabels = true;

// Set the row layout type to tabular
pt.Options.RowLayout = wasmModule.PivotTableLayoutType.Tabular;

// Set the Desc field as an additional row field
let r2 = pt.PivotFields.get_Item('Desc');
r2.Axis = wasmModule.AxisTypes.Row;

// Add the OnHand field as a data field with the label "Sum of onHand"
pt.DataFields.Add(pt.PivotFields.get_Item('OnHand'), 'Sum of onHand', wasmModule.SubtotalTypes.None);

// Set the built-in style for the pivot table appearance
pt.BuiltInStyle = wasmModule.PivotBuiltInStyles.PivotStyleMedium12;
```

---

# spire.xls javascript pivot table
## set format options for pivot table
```javascript
//Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });

let pt = sheet.PivotTables.get(0);
//Set the PivotTable report is automatically formatted
pt.Options.IsAutoFormat = true;

//Setting the PivotTable report shows grand totals for rows.
pt.ShowRowGrand = true;

//Setting the PivotTable report shows grand totals for columns.
pt.ShowColumnGrand = true;

//Setting the PivotTable report displays a custom string in cells that contain null values.
pt.DisplayNullString = true;
pt.NullString = 'null';

//Setting the PivotTable report's layout
pt.PageFieldOrder = wasmModule.PagesOrderType.DownThenOver;
```

---

# spire.xls javascript pivot table
## set pivot field format
```javascript
// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
// Access the first pivot table in the worksheet
let pt = sheet.PivotTables.get(0);
// Access the first pivot field in the pivot table
let pf = pt.PivotFields.get(0);

// Setting the field auto sort ascend.
pf.SortType = wasmModule.PivotFieldSortType.Ascending;

// Setting Subtotal auto show.
pf.SubtotalTop = true;

// Setting Subtotal as Count type
pf.Subtotals = wasmModule.SubtotalTypes.Count;

// Setting the field auto show.
pf.IsAutoShow = true;
```

---

# Spire.XLS JavaScript Pivot Table
## Show data field in row for pivot table
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
    fileName: excelFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Get the data in Pivot Table
let pivotTable = book.Worksheets.get(1).PivotTables.get(0);

// Show the datafield in row
pivotTable.ShowDataFieldInRow = true;

// Calculate Data
pivotTable.CalculateData();

// Save the workbook to the specified path
book.SaveToFile({
    fileName: outputFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Clean up resources
book.Dispose();
```

---

# Excel Pivot Table Subtotals
## Show subtotals in a pivot table
```javascript
// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'Pivot Table' });
// Get the first pivot table from the worksheet
let pt = sheet.PivotTables.get(0);

// Show Subtotals
pt.ShowSubtotals = true;
```

---

# Sorting Pivot Table in Excel
## This code demonstrates how to create and sort a pivot table in Excel using JavaScript
```javascript
// Add an empty worksheet
let sheet2 = book.CreateEmptySheet();
sheet2.Name = 'Pivot Table';

// Specify the data source
let dataRange = sheet.Range.get('A1:C9');
let cache = book.PivotCaches.Add({ range: dataRange });

// Add PivotTable
let pt = sheet2.PivotTables.Add('Pivot Table', sheet.Range.get('A1'), cache);

// Configure the pivot table settings
let r1 = pt.PivotFields.get_Item('No');
r1.Axis = wasmModule.AxisTypes.Row;
pt.Options.RowLayout = wasmModule.PivotTableLayoutType.Tabular;

// Sort the "No" field in descending order
r1.SortType = wasmModule.PivotFieldSortType.Descending;

let r2 = pt.PivotFields.get_Item('Name');
r2.Axis = wasmModule.AxisTypes.Row;
// Add a data field to the pivot table
pt.DataFields.Add(pt.PivotFields.get_Item('OnHand'), 'Sum of onHand', wasmModule.SubtotalTypes.None);
// Set the pivot table style
pt.BuiltInStyle = wasmModule.PivotBuiltInStyles.PivotStyleMedium12;
```

---

# Spire.XLS JavaScript Pivot Table
## Update Pivot Table Data Source
```javascript
// Create a new workbook
const book = wasmModule.Workbook.Create();
book.LoadFromFile({
    fileName: excelFileName,
    version: wasmModule.ExcelVersion.Version2010,
});

// Modify data of data source
let data = book.Worksheets.get('Data');
// Modify the data source by changing the value in cell A2 to "NewValue"
data.Range.get('A2').Text = 'NewValue';
// Modify the data source by changing the value in cell D2 to 28000
data.Range.get('D2').NumberValue = 28000;

// Get the sheet in which the pivot table is located
let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
// Get the first pivot table from the worksheet
let pt = sheet.PivotTables.get(0);

// Refresh and calculate
pt.Cache.IsRefreshOnLoad = true;
// Calculate and update the pivot table data
pt.CalculateData();
```

---

# Detect Excel Workbook Protection
## Check if an Excel workbook is password protected
```javascript
// Detect if the workbook is password protected
const value = wasmModule.Workbook.IsPasswordProtected(inputFileName);
let boolvalue = [];
boolvalue.push(value ? "Yes" : "No");
```

---

# Hide Formulas in Excel Worksheet
## Code to hide formulas and protect an Excel worksheet with password
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Hide the formulas in the used range
sheet.AllocatedRange.IsFormulaHidden = true;

// Protect the worksheet with password
sheet.Protect("e-iceblue");
```

---

# Excel Cell Locking
## Lock specific cells and protect worksheet with password
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();
// Create an empty worksheet
workbook.CreateEmptySheet();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Loop through all the rows in the worksheet and unlock them
for (let i = 0; i < 20; i++) {
  sheet.Rows.get(i).Style.Locked = false;
}

// Lock specific cell in the worksheet
sheet.Range.get("A1").Text = "Locked";
sheet.Range.get("A1").Style.Locked = true;

// Lock specific cell range in the worksheet
sheet.Range.get("C1:E3").Text = "Locked";
sheet.Range.get("C1:E3").Style.Locked = true;

// Set the password
sheet.Protect({
  password: "123",
  options: wasmModule.SheetProtectionType.All,
});
```

---

# Excel Column Locking
## Lock specific columns in a new Excel file with password protection
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();
// Create an empty worksheet
workbook.CreateEmptySheet();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Loop through all the columns in the worksheet and unlock them
for (let i = 0; i < 20; i++) {
  sheet.Rows.get(i).Style.Locked = false;
}

// Lock the fourth column in the worksheet
sheet.Columns.get(3).Text = "Locked";
sheet.Columns.get(3).Style.Locked = true;

// Set the password
sheet.Protect("123", wasmModule.SheetProtectionType.All);
```

---

# Lock Specific Row in Excel
## This code demonstrates how to lock a specific row in an Excel worksheet while keeping other rows unlocked
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
// Create an empty worksheet
workbook.CreateEmptySheet();

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Loop through all the rows in the worksheet and unlock them
for (let i = 0; i < 20; i++) {
  sheet.Rows.get(i).Style.Locked = false;
}

// Lock the third row in the worksheet
sheet.Rows.get(2).Text = "Locked";
sheet.Rows.get(2).Style.Locked = true;

// Set the password
sheet.Protect("123", wasmModule.SheetProtectionType.All);
```

---

# Excel Cell Protection
## Protect specific cells in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Protect cell
sheet.Range.get("B3").Style.Locked = true;
sheet.Range.get("C3").Style.Locked = false;

sheet.Protect("TestPassword", wasmModule.SheetProtectionType.All);
```

---

# Excel Protection with Editable Ranges
## Demonstrates how to protect an Excel worksheet while allowing specific ranges to remain editable
```javascript
// Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);

// Get the first worksheet
let sheet = workbook.Worksheets.get(0);

// Protect cell
let editableRange = sheet.Range.get("B4:E12");
let tname = "EditableRanges";
sheet.AddAllowEditRange({ title: tname, range: editableRange });

sheet.Protect("TestPassword", wasmModule.SheetProtectionType.All);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Workbook Protection
## Protect an Excel workbook with a password
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);

// Protect Workbook
workbook.Protect("e-iceblue");
```

---

# Excel Digital Signature Removal
## Remove all digital signatures from an Excel workbook
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);

//Remove all digital signatures
workbook.RemoveAllDigitalSignatures();
```

---

# Excel Worksheet Protection Unlock
## Unprotect a password-protected worksheet in Excel
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Unlock the worksheet with password
sheet.Unprotect({ password: "e-iceblue" });
```

---

# Spire.XLS JavaScript Security
## Unlock simple worksheet in Excel file
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);

//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Unlock the worksheet in an unlocked Excel file with null string
sheet.Unprotect();
```

---

# spire.xls javascript textbox
## extract text from textbox in excel
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile(inputFileName);

let sheet = workbook.Worksheets.get(0);

//Get the first textbox
let shape = sheet.TextBoxes.get(0);

//Extract text from the text box
let content = [];
content.push("The text extracted from the TextBox is: ");
content.push(shape.Text);
```

---

# spire.xls javascript textbox
## get TextBox by name in worksheet
```javascript
//Insert a TextBox
sheet.Range.get("A2").Text = "Name：";
let textBox = sheet.TextBoxes.AddTextBox(2, 2, 18, 65);

//Set the name
textBox.Name = "FirstTextBox";

//Set string text for TextBox
textBox.Text =
  "Spire.XLS for .NET is a professional Excel .NET component that can be used to any type of .NET 2.0, 3.5, 4.0 or 4.5 framework application, both ASP.NET web sites and Windows Forms application.";

//Get the TextBox by the name
let FindTextBox = sheet.TextBoxes.get("FirstTextBox");

//Get the TextBox text
let text = FindTextBox.Text;
```

---

# Excel Text Box Manipulation
## Modify text box content and alignment in Excel worksheets
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Get the first textbox
let tb = sheet.TextBoxes.get(0);

//Change the text of textbox
tb.Text = "Spire.XLS for .NET";

//Set the alignment of textbox as center
tb.HAlignment = wasmModule.CommentHAlignType.Center;
tb.VAlignment = wasmModule.CommentVAlignType.Center;
```

---

# spire.xls javascript textbox
## remove borderline of textbox in Excel chart
```javascript
// Create textbox1 in the chart and input text information
let textbox1 = chart.TextBoxes.AddTextBox(50, 50, 100, 600);
textbox1.Text = "The solution with borderline";

// Create textbox2 in the chart, input text information and remove borderline
let textbox2 = chart.TextBoxes.AddTextBox(1000, 50, 100, 600);
textbox2.Text = "The solution without borderline";
textbox2.Line.Weight = 0;
```

---

# Excel TextBox Text Replacement
## Replace text in textboxes within an Excel sheet
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);

const tag = "TAG_1$TAG_2";
const replace = "Spire.XLS for .NET$Spire.XLS for JAVA";

let tags = tag.split("$");
let replacements = replace.split("$");

for (let i = 0; i < tags.length; i++) {
  // Replace text in textbox
  _ReplaceTextInTextBox(sheet, `<${tags[i]}>`, replacements[i]);
}

function _ReplaceTextInTextBox(sheet, sFind, sReplace) {
  // Get the textboxes of sheet
  let textBoxes = sheet.TextBoxes;
  // replace text in each textbox
  for (let tb of textBoxes) {
    if (tb.Text && tb.Text.includes(sFind)) {
      tb.Text = tb.Text.replace(sFind, sReplace);
    }
  }
}
```

---

# spire.xls javascript textbox formatting
## set font and background color for textbox in excel
```javascript
//Get the textbox which will be edited.
let shape = sheet.TextBoxes.get(0);

//Set the font and background color for the textbox.
//Set font.
let font = workbook.CreateFont();
//font.IsStrikethrough = true
font.FontName = "Century Gothic";
font.Size = 10;
font.IsBold = true;
font.Color = wasmModule.Color.get_Blue();
let rto = shape.RichText;
let rt = wasmModule.RichTextShape.Convert(rto);
rt.SetFont(0, shape.Text.length - 1, font);

//Set background color
shape.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
shape.Fill.ForeKnownColor = wasmModule.ExcelColors.BlueGray;
```

---

# Excel Textbox Internal Margin Setting
## Set internal margins for textbox in Excel worksheet
```javascript
//Add a textbox to the sheet and set its position and size.
let textbox = sheet.TextBoxes.AddTextBox(4, 2, 100, 300);

//Set the text on the textbox.
textbox.Text = "Insert TextBox in Excel and set the margin for the text";
textbox.HAlignment = wasmModule.CommentHAlignType.Center;
textbox.VAlignment = wasmModule.CommentVAlignType.Center;

//Set the inner margins of the contents.
textbox.InnerLeftMargin = 1;
textbox.InnerRightMargin = 3;
textbox.InnerTopMargin = 1;
textbox.InnerBottomMargin = 1;
```

---

# spire.xls javascript textbox
## set wrap text for textbox
```javascript
//Get the text box
let shape = sheet.TextBoxes.get(0);

//Set wrap text
shape.IsWrapText = true;
```

---

# Activate Worksheet in Excel Workbook
## Core functionality to activate a specific worksheet in an Excel workbook
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({fileName: excelFileName});
//Get the second worksheet from the workbook
let sheet = workbook.Worksheets.get(1);

//Activate the sheet
sheet.Activate();
```

---

# Add Page Breaks in Excel
## Add horizontal and vertical page breaks to Excel worksheet
```javascript
// Add page break in Excel file.
sheet.HPageBreaks.Add(sheet.Range.get("E4"));
sheet.VPageBreaks.Add(sheet.Range.get("C4"));
```

---

# spire.xls javascript worksheet
## add worksheet to workbook
```javascript
//Add a new worksheet named AddedSheet
let sheet = workbook.Worksheets.Add("AddedSheet");
sheet.Range.get("C5").Text = "This is a new sheet.";
```

---

# Apply Style to Worksheet
## This code demonstrates how to create a style and apply it to an entire Excel worksheet
```javascript
//Create a workbook
let workbook = wasmModule.value.Workbook.Create();

//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Create a cell style
let style = workbook.Styles.Add("newStyle");
style.Color = wasmModule.value.Color.get_LightBlue();
style.Font.Color = wasmModule.value.Color.get_White();
style.Font.Size = 15;
style.Font.IsBold = true;

//Apply the style to the first worksheet
sheet.ApplyStyle(style);
```

---

# Check Dialog Sheet in Excel File
## This code checks if a worksheet in an Excel file is a dialog sheet
```javascript
// Get the sheet
let sheet = workbook.Worksheets.get(0);

//Check if the worksheet is a dialog sheet.
if (sheet.Type === spirexls.ExcelSheetType.DialogSheet) {
  // Worksheet is a Dialog Sheet
} else {
  // Worksheet is not a Dialog Sheet
}
```

---

# Copy Worksheet to Another Excel File
## This code demonstrates how to copy a worksheet from one Excel workbook to another
```javascript
// Create a workbook.
let workbook = wasmModule.Workbook.Create();
// Get the first worksheet.
let sheet = workbook.Worksheets.get(0);

// Create another Workbook.
let workbook1 = wasmModule.Workbook.Create();
// Get the first worksheet in the book.
let sheet1 = workbook1.Worksheets.get(0);
// Copy worksheet to destination worsheet in another Excel file.
sheet1.CopyFrom(sheet);
```

---

# Spire.XLS JavaScript Worksheet Copy
## Copy worksheet within workbook
```javascript
//Get the first worksheet and create a new worksheet.
let sheet = workbook.Worksheets.get(0);
let sheet1 = workbook.Worksheets.Add("MySheet");
let sourceRange = sheet.AllocatedRange;

//Copy the first worksheet to the second one.
sheet.Copy({
  sourceRange: sourceRange,
  worksheet: sheet1,
  destRow: sheet.FirstRow,
  destColumn: sheet.FirstColumn,
  copyStyle: true,
});
```

---

# Spire.XLS JavaScript Worksheets
## Copy only visible sheets from one workbook to another
```javascript
// Create a new workbook and clear the sheets
let workbookNew = wasmModule.Workbook.Create();
workbookNew.Version = spirexls.ExcelVersion.Version2013;
workbookNew.Worksheets.Clear();

// Loop through the worksheets
for (let sheet of workbook.Worksheets) {
  // Judge if the worksheet is visible
  if (sheet.Visibility === spirexls.WorksheetVisibility.Visible) {
    // Copy the sheet to new workbook
    let name = sheet.Name;
    workbookNew.Worksheets.AddCopy({ sheet: sheet });
  }
}
```

---

# Copy Worksheet Between Workbooks
## This code demonstrates how to copy a worksheet from one Excel workbook to another workbook using Spire.XLS for JavaScript.

```javascript
//Create a workbook and load a file
let sourceWorkbook = wasmModule.Workbook.Create();
sourceWorkbook.LoadFromFile({ fileName: excelFileName1 });
//Get the first worksheet
let srcWorksheet = sourceWorkbook.Worksheets.get(0);

//Create a workbook
let targetWorkbook = wasmModule.Workbook.Create();

//Load the target Excel document from disk
targetWorkbook.LoadFromFile({ fileName: excelFileName2 });

//Add a new worksheet
let targetWorksheet = targetWorkbook.Worksheets.Add("added");

//Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
targetWorksheet.CopyFrom(srcWorksheet);

const outputFileName = "CopyWorksheet_output.xlsx";
targetWorkbook.SaveToFile({ fileName: outputFileName });
```

---

# Detect Empty Worksheet
## Check if worksheets in an Excel workbook are empty
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({ fileName: excelFileName });

let worksheet1 = workbook.Worksheets.get(0);

//Detect the first worksheet is empty or not
let detect1 = worksheet1.IsEmpty;

//Get the second worksheet
let worksheet2 = workbook.Worksheets.get(1);

//Detect the second worksheet is empty or not
let detect2 = worksheet2.IsEmpty;

//Set string format for displaying
let result = `The first worksheet is empty or not: ${detect1}\r\nThe second worksheet is empty or not: ${detect2}`;
```

---

# spire.xls javascript worksheet
## fill data in worksheet cells
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Get first worksheet of the workbook
let worksheet = workbook.Worksheets.get(0);

// Fill data
worksheet.Range.get("A1").Style.Font.IsBold = true;
worksheet.Range.get("B1").Style.Font.IsBold = true;
worksheet.Range.get("C1").Style.Font.IsBold = true;
worksheet.Range.get("A1").Text = "Month";
worksheet.Range.get("A2").Text = "January";
worksheet.Range.get("A3").Text = "February";
worksheet.Range.get("A4").Text = "March";
worksheet.Range.get("A5").Text = "April";
worksheet.Range.get("B1").Text = "Payments";
worksheet.Range.get("B2").NumberValue = 251;
worksheet.Range.get("B3").NumberValue = 515;
worksheet.Range.get("B4").NumberValue = 454;
worksheet.Range.get("B5").NumberValue = 874;
worksheet.Range.get("C1").Text = "Sample";
worksheet.Range.get("C2").Text = "Sample1";
worksheet.Range.get("C3").Text = "Sample2";
worksheet.Range.get("C4").Text = "Sample3";
worksheet.Range.get("C5").Text = "Sample4";

// Set width for the second column
worksheet.SetColumnWidth(2, 10);
```

---

# Excel Worksheet Freeze Panes
## Freeze top row in Excel worksheet using JavaScript
```javascript
//Get the first sheet
let sheet = workbook.Worksheets.get(0);
//Freeze Top Row
sheet.FreezePanes(2, 1);

//Set width for the second column
sheet.SetColumnWidth(2, 10);
```

---

# Get Freeze Pane Range in Excel Worksheet
## Retrieve the row and column index of freeze panes in an Excel worksheet
```javascript
// Get the first sheet
let sheet = workbook.Worksheets.get(0);
let rowIndex = null;
let colIndex = null;
let r = [];
//The row and column index of the frozen pane is passed through the out parameter.
//If it returns to 0, it means that it is not frozen
let indexs = sheet.GetFreezePanes();
colIndex = indexs[1];
rowIndex = indexs[0];

r.push(`Row index: ${rowIndex}, column index: ${colIndex}`);
```

---

# Excel Font List Extraction
## Get list of fonts used in Excel workbook
```javascript
// Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({ fileName: excelFileName });

let fonts = [];

// Loop all sheets of workbook
for (let i = 0; i < workbook.Worksheets.Count; i++) {
  let sheet = workbook.Worksheets.get(i);
  for (let r = 0; r < sheet.Rows.Count; r++) {
    for (let c = 0; c < sheet.Rows.get(r).Cells.Count; c++) {
      // Get the font of cell and add it to list
      let cell = sheet.Rows.get(r).Cells.get(c);
      fonts.push(cell.Style.Font);
    }
  }
}

let strB = [];

for (let font of fonts) {
  strB.push(`FontName:${font.FontName}; FontSize:${font.Size}`);
}
```

---

# spire.xls javascript worksheet
## get paper size of worksheets
```javascript
// loop the worksheet and get the PageSetup of sheet
let sb = [];
for (let i = 0; i < workbook.Worksheets.Count; i++) {
  let sheet = workbook.Worksheets.get(i);
  let width = sheet.PageSetup.PageWidth;
  let height = sheet.PageSetup.PageHeight;
  sb.push(sheet.Name);
  sb.push(`Width: ${width}\tHeight: ${height}\r\n`);
}
```

---

# Get Worksheet Names
## Extract names of all worksheets in an Excel workbook
```javascript
// Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({ fileName: excelFileName });

// get each name of worksheet
let sb = [];
for (let i = 0; i < workbook.Worksheets.Count; i++) {
  let sheet = workbook.Worksheets.get(i);
  sb.push(sheet.Name);
}
```

---

# spire.xls javascript worksheet
## hide or show worksheet
```javascript
// Hide the sheet named "Sheet1"
let sheet1 = workbook.Worksheets.get("Sheet1");
sheet1.Visibility = spirexls.WorksheetVisibility.Hidden;

// Show the second sheet
let sheet2 = workbook.Worksheets.get(1);
sheet2.Visibility = spirexls.WorksheetVisibility.Visible;
```

---

# spire.xls javascript worksheet
## hide worksheet tabs
```javascript
// Hide worksheet tab
workbook.ShowTabs = false;
```

---

# Excel Worksheet Hide Zero Values
## Hide zero values in Excel worksheet using JavaScript
```javascript
// Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({ fileName: excelFileName });

// Get the first sheet
let sheet = workbook.Worksheets.get(0);
// Set false to hide the zero values in sheet
sheet.IsDisplayZeros = false;

// Save the workbook
workbook.SaveToFile({ fileName: outputFileName });

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Document Property Linking
## Link custom document property to content in Excel workbook
```javascript
// Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({ fileName: excelFileName });

// Add a custom document property
workbook.CustomDocumentProperties.Add("Test", "MyNamedRange");
// Get the added document property
let properties = workbook.CustomDocumentProperties;
let property = properties.get("Test");
// Link to content
property.LinkToContent = true;
```

---

# Excel Worksheet Movement
## Move worksheet to a different position in Excel workbook
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);
// Move worksheet
sheet.MoveWorksheet(2);
```

---

# Excel Page Break Preview
## Set zoom scale for PageBreakView mode in Excel worksheet
```javascript
// Get the first worksheet
let sheet = workbook.Worksheets.get(0);
// Set the scale of PageBreakView mode in Excel file
sheet.ZoomScalePageBreakView = 80;
```

---

# Removing Page Breaks in Excel Worksheet
## This code demonstrates how to remove vertical and horizontal page breaks in an Excel worksheet
```javascript
// Get the first worksheet from the workbook
let sheet = workbook.Worksheets.get(0);

// Clear all the vertical page breaks
sheet.VPageBreaks.Clear();

// Remove the first horizontal Page Break
sheet.HPageBreaks.RemoveAt(0);

// Set the ViewMode as Preview to see how the page breaks work
sheet.ViewMode = spirexls.ViewMode.Preview;
```

---

# Spire.XLS JavaScript Worksheet Removal
## Remove a worksheet from an Excel workbook by index
```javascript
// Create a workbook
const workbook = wasmModule.Workbook.Create();

// Remove a worksheet by sheet index
workbook.Worksheets.RemoveAt(1);

// Dispose of the workbook object to release resources
workbook.Dispose();
```

---

# Excel Page Break Setting
## Set horizontal and vertical page breaks in Excel worksheet
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Set Excel Page Break Horizontally
sheet.HPageBreaks.Add(sheet.Range.get("A8"));
sheet.HPageBreaks.Add(sheet.Range.get("A14"));

//Set Excel Page Break Vertically
//sheet.VPageBreaks.Add(sheet.Range.get("B1"));
//sheet.VPageBreaks.Add(sheet.Range.get("C1"));

//Set view mode to Preview mode
workbook.Worksheets.get(0).ViewMode = spirexls.ViewMode.Preview;
```

---

# spire.xls javascript worksheet
## set worksheet tab color
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();

//Set the tab color of first sheet to be red
let worksheet = workbook.Worksheets.get(0);
worksheet.TabColor = wasmModule.Color.get_Red();

//Set the tab color of second sheet to be green
worksheet = workbook.Worksheets.get(1);
worksheet.TabColor = wasmModule.Color.get_Green();

//Set the tab color of third sheet to be blue
worksheet = workbook.Worksheets.get(2);
worksheet.TabColor = wasmModule.Color.get_LightBlue();
```

---

# Worksheet View Mode Setting
## Set worksheet view mode to preview
```javascript
//Set the view mode
workbook.Worksheets.get(0).ViewMode = spirexls.ViewMode.Preview;
```

---

# Spire.XLS JavaScript Grid Lines Control
## Show or hide grid lines in Excel worksheets
```javascript
// Get the first and second worksheet
let sheet1 = workbook.Worksheets.get(0);
let sheet2 = workbook.Worksheets.get(1);

// Hide grid line in the first worksheet
sheet1.GridLinesVisible = false;
// Show grid line in the second worksheet
sheet2.GridLinesVisible = true;
```

---

# spire.xls javascript worksheet
## show worksheet tabs in excel workbook
```javascript
//Show worksheet tab
workbook.ShowTabs = true;
```

---

# spire.xls javascript worksheet panes
## split excel worksheet into multiple panes
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);

//Vertical and horizontal split the worksheet into four panes
sheet.FirstVisibleColumn = 2;
sheet.FirstVisibleRow = 5;
sheet.VerticalSplit = 4000;
sheet.HorizontalSplit = 5000;

//Set the active pane
sheet.ActivePane = 1;
```

---

# spire.xls javascript worksheet
## unfreeze excel panes
```javascript
//Create a workbook
const workbook = wasmModule.Workbook.Create();
//Get the first worksheet.
let sheet = workbook.Worksheets.get(0);
//Unfreeze the panes.
sheet.RemovePanes();
```

---

# Worksheet Password Verification
## Verify if a worksheet is password protected
```javascript
// Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();
workbook.LoadFromFile({ fileName: excelFileName });

// Get the first worksheet
let worksheet = workbook.Worksheets.get(0);

// Verify the first worksheet
let detect = worksheet.IsPasswordProtected;
```

---

# spire.xls javascript zoom factor
## set worksheet zoom factor
```javascript
//Get the first worksheet
let sheet = workbook.Worksheets.get(0);
//Set the zoom factor of the sheet to 85
sheet.Zoom = 85;
```

---

# Accessing Document Properties in Excel Workbook
## Demonstrates how to access document properties by name and index in an Excel workbook
```javascript
// Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile({ fileName: excelFileName });

// Get all document properties
let properties = book.CustomDocumentProperties;

// Access document property by property name
let property1 = properties.get({ strName: "Editor" });
let obj = spirexls.String.Convert(property1.Value);
builder.push(`${property1.Name} ${obj.Value}`);

// Access document property by property index
let property2 = properties.get({ iIndex: 0 });
let obj2 = spirexls.String.Convert(property2.Value).Value;
builder.push(`${property2.Name} ${obj2}`);

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# Spire.XLS JavaScript Custom Properties
## Add custom document properties to workbook
```javascript
//Add a custom property to make the document as final
workbook.CustomDocumentProperties.Add({
  strName: "_MarkAsFinal",
  boolValue: true,
});

//Add other custom properties to the workbook
workbook.CustomDocumentProperties.Add("The Editor", "E-iceblue");
workbook.CustomDocumentProperties.Add({
  strName: "Phone number",
  intValue: 81705109,
});
workbook.CustomDocumentProperties.Add({
  strName: "Revision number",
  dblValue: 7.12,
});
workbook.CustomDocumentProperties.Add({
  strName: "Revision date",
  objValue: wasmModule.DateTime.get_Now(),
});
```

---

# spire.xls javascript workbook decryption
## remove password protection from Excel workbook
```javascript
//Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.OpenPassword = "eiceblue";
book.LoadFromFile(excelFileName);

//Decrypt workbook
book.UnProtect();

//Save the document
book.SaveToFile({ fileName: outputFileName });

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# Excel Version Detection
## Detect the version of an Excel file using Spire.XLS for JavaScript
```javascript
// Create a workbook and load an Excel file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

// Get the version of the Excel file
const version = book.Version;
```

---

# Excel VBA Macros Detection
## Detect if an Excel document contains VBA macros
```javascript
//Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

//Detect if the Excel file contains VBA macros
let hasMacros = book.HasMacros;
```

---

# spire.xls javascript encryption
## encrypt workbook with password
```javascript
// Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

// Protect Workbook with the password you want
book.Protect("eiceblue");
```

---

# spire.xls javascript workbook properties
## retrieve Excel document properties and custom properties
```javascript
// Get the general excel properties
let properties1 = workbook.DocumentProperties;
let sb = [];
sb.push("Excel Properties:");
for (let i = 0; i < properties1.Count; i++) {
  let name = properties1.get(i).Name;
  let obj = properties1.get(i).Value;
  let t = properties1.get(i).PropertyType;
  let value = null;
  if (t === wasmModule.PropertyType.Double) {
    value = wasmModule.Double.Convert(obj).Value;
  } else if (t === wasmModule.PropertyType.DateTime) {
    value = wasmModule.DateTime.Convert(obj).ToString();
  } else if (t === wasmModule.PropertyType.Bool) {
    value = wasmModule.Boolean.Convert(obj).Value;
  } else if (
    t === wasmModule.PropertyType.Int ||
    t === wasmModule.PropertyType.Int32
  ) {
    value = wasmModule.Int32.Convert(obj).Value;
  } else {
    value = wasmModule.String.Convert(obj).Value;
  }
  sb.push(name + ": " + String(value));
}
sb.push("");

// Get the custom properties
let properties2 = workbook.CustomDocumentProperties;
sb.push("Custom Properties:");
for (let i = 0; i < properties2.Count; i++) {
  let name = properties2.get(i).Name;
  let t = properties2.get(i).PropertyType;
  let obj = properties2.get(i).Value;
  let value = null;
  if (t === wasmModule.PropertyType.Double) {
    value = wasmModule.Double.Convert(obj).Value;
  } else if (t === wasmModule.PropertyType.DateTime) {
    value = wasmModule.DateTime.Convert(obj).ToString();
  } else if (t === wasmModule.PropertyType.Bool) {
    value = wasmModule.Boolean.Convert(obj).Value;
  } else if (
    t === wasmModule.PropertyType.Int ||
    t === wasmModule.PropertyType.Int32
  ) {
    value = wasmModule.Int32.Convert(obj).Value;
  } else {
    value = wasmModule.String.Convert(obj).Value;
  }
  sb.push(name + ": " + String(value));
}
```

---

# Excel Workbook Window Hiding
## Demonstrates how to hide the window of an Excel workbook using Spire.XLS for JavaScript
```javascript
// Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

// Hide window
book.IsHideWindow = true;

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# Excel Workbook with Macro Handling
## Load and save Excel files with macros
```javascript
// Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile("MacroSample.xls");

// Get the first worksheet
let sheet = book.Worksheets.get(0);
// Set value for cell A5
sheet.Range.get("A5").Text = "This is a simple test!";

// Save the document
book.SaveToFile({ fileName: "LoadAndSaveFileWithMacro_output.xlsx" });

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# Spire.XLS JavaScript Workbook Merge
## Merge multiple Excel files into a single workbook
```javascript
// Create a new workbook
let newbook = wasmModule.Workbook.Create();
newbook.Version = wasmModule.ExcelVersion.Version2013;
// Clear all worksheets
newbook.Worksheets.Clear();

let tempbook = wasmModule.Workbook.Create();

// Files to merge
const files = [
  "MergeExcelFiles-1.xlsx",
  "MergeExcelFiles-2.xls",
  "MergeExcelFiles-3.xlsx",
];

for (const file of files) {
  // Load the file
  tempbook.LoadFromFile(file.split("/").pop());

  for (let i = 0; i < tempbook.Worksheets.Count; i++) {
    let sheet = tempbook.Worksheets.get(i);
    // Copy every sheet in a workbook
    wasmModule.XlsWorksheetsCollection.Convert(
      newbook.Worksheets
    ).AddCopy({
      sheet: sheet,
      flags: wasmModule.WorksheetCopyType.CopyAll,
    });
  }
}

// Save the merged document
newbook.SaveToFile({ fileName: "MergeExcelFiles_output.xlsx" });
```

---

# Open Encrypted Excel File
## Demonstrates how to try different passwords to open an encrypted Excel file
```javascript
// Create string builder
let builder = [];

const passwords = ["password1", "password2", "password3", "1234"];
for (let i = 0; i < passwords.length; i++) {
  try {
    // Create a workbook
    let workbook = wasmModule.Workbook.Create();

    // Open password
    workbook.OpenPassword = passwords[i];

    // Load the document
    workbook.LoadFromFile(excelFileName);

    builder.push(
      "Password = " +
        passwords[i] +
        " is correct. The encrypted Excel file opened successfully!"
    );
  } catch (e) {
    builder.push("Password = " + passwords[i] + " is not correct");
    builder.push("ErrorMessage = " + e.message); // Capture exception message
  }
}
```

---

# Excel File Opening in JavaScript
## Demonstrates various methods to open Excel files using Spire.XLS for JavaScript
```javascript
// 1. Load file by file path
// Create a workbook
let workbook1 = wasmModule.Workbook.Create();
// Load the document from disk
workbook1.LoadFromFile({ fileName: inputFile });

// 2. Load file by file stream
let stream = wasmModule.Stream.CreateByFile(inputFile.split("/").pop());
// Create a workbook
let workbook2 = wasmModule.Workbook.Create();
// Load the document from stream
workbook2.LoadFromStream(stream);
stream.Close();

// 3. Open Microsoft Excel 97 - 2003 file
let wbExcel97 = wasmModule.Workbook.Create();
wbExcel97.LoadFromFile({ fileName: inputFile_97 });

// 4. Open xml file
let wbXML = wasmModule.Workbook.Create();
wbXML.LoadFromXml(inputFile_xml);

// 5. Open csv file
let wbCSV = wasmModule.Workbook.Create();
wbCSV.LoadFromFile({
  fileName: inputFile_csv,
  separator: ",",
  row: 1,
  column: 1,
});
```

---

# Reading Excel from Stream
## Demonstrates how to load an Excel workbook from a stream
```javascript
//Create a workbook and load a file
const workbook = wasmModule.Workbook.Create();

// Open excel from a stream
let fileStream = wasmModule.Stream.CreateByFile(excelFileName);
workbook.LoadFromStream(fileStream);
```

---

# spire.xls javascript workbook
## remove custom properties from Excel workbook
```javascript
// Retrieve a list of all custom document properties of the Excel file
let customDocumentProperties = book.CustomDocumentProperties;

// Remove "Editor" custom document property
customDocumentProperties.Remove("Editor");
```

---

# Excel File Conversion
## Save Excel workbook to various file formats
```javascript
//Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

//Save to xls
let output_xls = "SaveFiles_output.xls";
book.SaveToFile({ fileName: outputDirectoryName + output_xls });

//Save to xlsx
let output_xlsx = "SaveFiles_output.xlsx";
book.SaveToFile({ fileName: outputDirectoryName + output_xlsx });

//Save to xlsb
let output_xlsb = "SaveFiles_output.xlsb";
book.SaveToFile({ fileName: outputDirectoryName + output_xlsb });

//Save to ods
let output_ods = "SaveFiles_output.ods";
book.SaveToFile({ fileName: outputDirectoryName + output_ods });

//Save to pdf
let output_pdf = "SaveFiles_output.pdf";
book.SaveToFile({ fileName: outputDirectoryName + output_pdf });

//Save to xml
let output_xml = "SaveFiles_output.xml";
book.SaveToFile({ fileName: outputDirectoryName + output_xml });

//Save to xps
let output_xps = "SaveFiles_output.xps";
book.SaveToFile({ fileName: outputDirectoryName + output_xps });

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# Excel Workbook Save to Stream
## Save an Excel workbook to a stream using JavaScript
```javascript
//Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

let outputFileName = "SaveStream_output.xlsx";
//Save the document to stream
let fileStream = wasmModule.Stream.CreateByFile(outputFileName);
book.SaveToStream(fileStream, wasmModule.FileFormat.Version2010);

// Dispose of the workbook object to release resources
book.Dispose();
```

---

# Setting Excel Calculation Mode
## Set the calculation mode of an Excel workbook to Manual
```javascript
//Create a workbook and load a file
const book = wasmModule.Workbook.Create();
book.LoadFromFile(excelFileName);

// Set excel calculation mode as Manual
book.CalculationMode = spirexls.ExcelCalculationMode.Manual;
```

---

# spire.xls javascript page margins
## set worksheet page margins
```javascript
// Get the first worksheet
let sheet = book.Worksheets.get(0);

// Set margins for top, bottom, left and right, here the unit of measure is Inch
sheet.PageSetup.TopMargin = 0.3;
sheet.PageSetup.BottomMargin = 1;
sheet.PageSetup.LeftMargin = 0.2;
sheet.PageSetup.RightMargin = 1;
// Set the header margin and footer margin
sheet.PageSetup.HeaderMarginInch = 0.1;
sheet.PageSetup.FooterMarginInch = 0.5;
```

---

# spire.xls javascript theme
## set workbook theme color
```javascript
// Create a workbook
let srcWorkbook = wasmModule.Workbook.Create();
// Load an excel file
srcWorkbook.LoadFromFile(inputFileName);
let srcWorksheet = srcWorkbook.Worksheets.get(0);

let workbook = wasmModule.Workbook.Create();
workbook.Worksheets.Clear();
workbook.Worksheets.AddCopy({ sheet: srcWorksheet });

// 1. Copy the theme of the workbook
// workbook.CopyTheme(srcWorkbook);

// 2. Set a certain type of color of the default theme in the workbook
workbook.SetThemeColor(
  wasmModule.ThemeColorType.Dk1,
  wasmModule.Color.get_SkyBlue()
);
```

---

# Excel Track Changes Management
## Accept or reject all tracked changes in an Excel workbook
```javascript
// Create a workbook and load a file
const book = wasmModule.Workbook.Create();        
book.LoadFromFile(excelFileName);

// Accept the changes or reject the changes
// workbook.AcceptAllTrackedChanges();
book.RejectAllTrackedChanges();
```

---

