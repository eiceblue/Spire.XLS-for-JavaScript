<template>
  <span>Click the following button to remove picture border</span>
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
        let inputFileName='RemovePictureBorder.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        // Get the first worksheet
        let sheet1 = workbook.Worksheets.get(0);

        // Get the first picture from the first worksheet
        let picture = sheet1.Pictures.get(0);

        // Remove the picture border
        // Method-1:
        picture.Line.Visible = false;

        // Method-2:
        // picture.Line.Weight = 0
 
        const outputFileName = 'RemovePictureBorder-out.xlsx';
        // Save the modified workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();
        
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

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