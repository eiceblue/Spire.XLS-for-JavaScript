<template>
  <span>Click the following button to set the alignment of text in the cell</span>
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

        // Input file
        let excelFileName='TextAlign.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        // Get the first worksheet
        const sheet = workbook.Worksheets.get(0);

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

        //Save result file
        const outputFileName = 'TextAlign_out.xlsx';
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
