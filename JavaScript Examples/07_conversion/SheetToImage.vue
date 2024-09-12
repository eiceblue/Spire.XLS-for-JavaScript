<template>
  <span>Click the following button to convert worksheet to image </span>
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

        let inputFileName='SheetToImage.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Convert the sheet to image and save it
        let image = sheet.ToImage(sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn);

        const outputFileName ='SheetToImage-out.png';
        // Save image to file
        image.Save(outputFileName);
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'image/png'});

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
