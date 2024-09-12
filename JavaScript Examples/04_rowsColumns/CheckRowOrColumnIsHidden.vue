<template>
  <span>Click the following button to check whether a row or column is hidden</span>
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
        let excelFileName='CheckRowOrColumnIsHidden.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});
        
        // Get the first worksheet in the workbook
        let sheet = workbook.Worksheets.get(0);
        let result = [];
        let rowIndex = 2;
        let columnIndex = 2;
        let rowIsHide = sheet.GetRowIsHide(rowIndex);
        if (rowIsHide) {
          result.push("The second row is hidden.");
        } else {
          result.push("The second row is not hidden.");
        }
        let columnIsHide = sheet.GetColumnIsHide(columnIndex);
        if (columnIsHide) {
          result.push("The second column is hidden.");
        } else {
          result.push("The second column is not hidden.");
        }

        //Save result file
        const outputFileName = 'CheckRowOrColumnIsHidden_out.txt';
        FS.writeFile(outputFileName, result.join("\n"))

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'text/plain'});

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
