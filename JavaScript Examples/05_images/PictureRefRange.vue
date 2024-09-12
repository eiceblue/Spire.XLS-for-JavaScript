<template>
  <span>Click the following button to set the reference range of a picture</span>
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
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='PictureRefRange.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);
        // Set values in cells A1 and B3.
        sheet.Range.get("A1").Value = "Spire.XLS";
        sheet.Range.get("B3").Value = "E-iceblue";

        // Get the first picture in worksheet
        let picture = sheet.Pictures.get(0);

        // Set the reference range of the picture to A1:B3
        picture.RefRange = "A1:B3";
 
        const outputFileName = 'PictureRefRange-out.xlsx';
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