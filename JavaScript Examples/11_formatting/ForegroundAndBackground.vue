<template>
  <span>Click the following button to set foreground and background</span>
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
      
        workbook.Version = wasmModule.ExcelVersion.Version2010;
      
        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);
      
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

        //Save result file
        const outputFileName = 'ForegroundAndBackground_out.xlsx';
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
