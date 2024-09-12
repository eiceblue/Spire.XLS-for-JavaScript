<template>
  <span>Click the following button to use style object</span>
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
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

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

        //Save result file
        const outputFileName = 'UsingStyleObject_out.xlsx';
        workbook.SaveToFile({fileName: outputFileName, version:wasmModule.ExcelVersion.Version2010});

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        // Download the result file
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
