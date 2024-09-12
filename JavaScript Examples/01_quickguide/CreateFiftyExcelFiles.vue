<template>
  <span>Click the following button to create fifty Excel files in one time</span>
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

        let outputFileName;

        // Loop to create 50 Excel workbooks, each containing 5 sheets
        for (let n = 0; n < 50; n++) {
          // Create a new workbook
          let workbook = wasmModule.Workbook.Create();
          // Add 5 empty sheets to the workbook
          workbook.CreateEmptySheets(5);
          // Fill the worksheets with data
          for (let i = 0; i < 5; i++) {
              let sheet = workbook.Worksheets.get(i);
              sheet.Name = `Sheet${i}`;
              for (let row = 1; row <= 151; row++) {
                  for (let col = 1; col <= 51; col++) {
                      sheet.Range.get({row:row, column:col}).Text = `row${row} col${col}`;
                  }
              }
          }

          // Define the output file name 
          outputFileName = `CreateFiftyExcelFiles_${n + 1}.xlsx`;

          // Save the workbook to the specified path
          workbook.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.ExcelVersion.Version2010});
        }
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
