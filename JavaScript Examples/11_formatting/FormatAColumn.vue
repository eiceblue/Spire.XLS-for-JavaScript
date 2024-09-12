<template>
  <span>Click the following button to format a column</span>
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

        //Save result file
        const outputFileName = 'FormatAColumn_out.xlsx';
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
