<template>
  <span>Click the following button to reset size and position for image</span>
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
        let inputFileName='SpireXls.png';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new workbook 
        const workbook = wasmModule.Workbook.Create();

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Add a picture to the first worksheet
        let picture = sheet.Pictures.Add({topRow:1, leftColumn:1, fileName:inputFileName});

        // Set the size for the picture
        picture.Width = 200;
        picture.Height = 200;

        // Set the position for the picture
        picture.Left = 200;
        picture.Top = 100;
 
        const outputFileName = 'ResetSizeAndPositionForImage-out.xlsx';
        // Save the modified workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
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