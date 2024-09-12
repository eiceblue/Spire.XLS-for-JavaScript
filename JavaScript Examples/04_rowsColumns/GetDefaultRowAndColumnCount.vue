<template>
  <span>Click the following button to get default row and column count of worksheet</span>
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

        //Save result file
        const outputFileName = 'GetDefaultRowAndColumnCount_out.txt';
		    FS.writeFile(outputFileName,sb.join("\n"))

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
