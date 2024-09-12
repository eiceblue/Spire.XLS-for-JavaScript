# Spire.XLS-for-JavaScript
A professional Excel development component that can be used to create, read, write, and convert Excel files in web applications with JavaScript.

[![Foo](https://i.imgur.com/L6PhXkQ.png)](https://www.e-iceblue.com/Introduce/xls-for-javascript.html)

[Product Page](https://www.e-iceblue.com/Introduce/xls-for-javascript.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-xls-f4.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)

[Spire.XLS for JavaScript](https://www.e-iceblue.com/Introduce/xls-for-javascript.html) is a powerful Excel JavaScript library that can be used to create, read, write, and convert Excel files in any JavaScript environment with no reliance on Microsoft Office Excel. This JavaScript library supports both client-side and server-side development, including environments like Node.js, and enables smooth integration with frameworks such as Vue, React, Angular, and pure JavaScript.

The API supports old Excel 97-2003 formats (.xls) and modern Excel formats like Excel 2007, Excel 2010, Excel 2013, Excel 2016, and Excel 2019 (.xlsx, .xlsb, .xlsm), as well as OpenOffice (.ods). It offers fast performance and reliability, reducing the complexity of manual Excel manipulation and avoiding the need for Microsoft Automation.

### 100% Standalone JavaScript API
Spire.XLS for JavaScript is a completely standalone Excel manipulation library that does not require Microsoft Excel or Office to be installed on the system.

### Freely Operate Excel Files
- Create/Save/Merge/Split/Get Excel files.
- Protect/Encrypt/Decrypt Excel files.
- Create/Add/Rename/Edit/Delete/Move worksheets.
- Insert/Modify/Remove hyperlinks.
- Add/Remove/Change/Hide/Show comments in Excel.
- Merge/Unmerge Excel cells, freeze/unfreeze panes, and insert/delete rows and columns.
- Add/Read/Calculate/Remove Excel formulas.
- Create/Refresh pivot tables.
- Apply/Remove conditional formatting.
- Add/Set/Change headers and footers.

### Easily Manipulate Cells & Excel Calculation Engine at Runtime
Developers can easily manipulate Excel cells and evaluate formulas in JavaScript at runtime. The fast, scalable calculation engine is compatible with Excel versions from 97-2003 to 2019. The library supports a wide range of cell formatting options, including cell merging, text wrapping, alignment, rotation, interior, borders, and font formatting (e.g., font type, size, color, bold, italic, strikeout, underline). Conditional formatting, search/replace, filtering, and data validation are also supported.

### Powerful & High-Quality Excel File Conversion
- Convert Excel to PDF/HTML/XML/CSV/Image/XPS/SVG.
- Convert CSV to Excel/CSV to PDF.
- Convert a selected range of cells to PDF.
- Convert XLS to XLSX or XSLX to XLS.
- Convert Excel to OpenDocument Spreadsheet (.ods).
- Save Excel charts as SVG/Image.
- Convert HTML to Excel.

### Chart, Data, and Other Elements
Spire.XLS for JavaScript provides a variety of chart types such as Pie Chart, Bar Chart, Column Chart, Line Chart, and Radar Chart. It supports seamless data transportation between databases and Excel in JavaScript. Hyperlinks and templates are also supported, making it easy to integrate Excel functionality into your web applications.

## Vue Examples

### Create an Excel File in JavaScript
```JavaScript
<template>
  <span>Click the following button to create my first Excel</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

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

        // Define the output file name 
        const outputFileName = 'HelloWorld.xlsx';

        // Save the workbook to the specified path
        workbook.SaveToFile({fileName: outputFileName, version: wasmModule.ExcelVersion.Version2010});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
        
        // Clean up resources
        workbook.Dispose();
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
```

### Convert Excel to PDF in JavaScript
```JavaScript
<template>
  <span>Click the following button to convert Excel to PDF</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      if (wasmModule) {
        
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='ToPDF.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        const outputFileName = 'ToPDF-out.pdf';
        //Save to PDF
        workbook.SaveToFile({fileName: outputFileName , fileFormat: wasmModule.FileFormat.PDF});
        // Dispose of the object to release resources
        workbook.Dispose();
        
        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/pdf'});

        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
```

### Convert Excel to Image in JavaScript
```JavaScript
<template>
  <span>Click the following button to convert worksheet to image </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='SheetToImage.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Convert the sheet to image and save it
        let image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);

        const outputFileName ='SheetToImage-out.png';
        // Save image to file
        image.Save(outputFileName);
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'image/png'});

        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
```

[Product Page](https://www.e-iceblue.com/Introduce/xls-for-javascript.html) | Documentation | Examples | [Forum](https://www.e-iceblue.com/forum/spire-xls-f4.html) | [Temporary License](https://www.e-iceblue.com/TemLicense.html) | [Customized Demo](https://www.e-iceblue.com/Misc/customized-demo.html)
