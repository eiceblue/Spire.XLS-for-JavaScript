<template>
  <span>Click the following button to create my first Excel</span>
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

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Clear default worksheets
        workbook.Worksheets.Clear();

        // Add a new worksheet named "MySheet"
        const sheet = workbook.Worksheets.Add("MySheet");

        // Set text for the "A1" range
        sheet.Range.get("A1").Text = "Hello World";

        // Set the column width to auto fit
        sheet.Range.get("A1").AutoFitColumns();

        // Define the output file name 
        const outputFileName = 'HelloWorld.xlsx';

        // Save the workbook to the specified path
        workbook.SaveToFile({fileName: outputFileName, version: wasmModule.ExcelVersion.Version2010});

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
