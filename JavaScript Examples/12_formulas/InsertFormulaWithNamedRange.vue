<template>
  <span>Click the following button to insert formula with named range </span>
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
        const sheet = workbook.Worksheets.get(0);

        //Set value
        sheet.Range.get("A1").Value = "1";
        sheet.Range.get("A2").Value = "1";

        //Create a named range
        const NamedRange = workbook.NameRanges.Add("NewNamedRange");
        NamedRange.NameLocal = "=SUM(A1+A2)";

        //Set the formula
        sheet.Range.get("C1").Formula = "NewNamedRange";

        //Save result file
        const outputFileName = 'InsertFormulaWithNamedRange_out.xlsx';
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
