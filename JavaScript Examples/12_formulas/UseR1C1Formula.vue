<template>
  <span>Click the following button to use R1C1 formula</span>
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

        // Get the first sheet
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

        // Write array formula
        sheet.Range.get("A5:C6").FormulaArray = "=LINEST(A1:A3,B1:C3,TRUE,TRUE)";

        // Calculate Formulas
        workbook.CalculateAllValue();

        //Save result file
        const outputFileName = 'UseR1C1Formula_out.xlsx';
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
