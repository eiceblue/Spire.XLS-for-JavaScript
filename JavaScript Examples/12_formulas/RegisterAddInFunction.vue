<template>
  <span>Click the following button to register AddIn function</span>
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
        let fileName='Test.xlam';
        await wasmModule.FetchFileToVFS(fileName, '', `${import.meta.env.BASE_URL}static/data/`);

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

        //Save result file
        const outputFileName = 'RegisterAddInFunction_out.xlsx';
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
