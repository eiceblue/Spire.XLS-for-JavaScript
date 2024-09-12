<template>
  <span>Click the following button to convert specific worksheet cells to image </span>
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

        let inputFileName='SpecificCellsToImage.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        
        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        const outputFileName ='SpecificCellsToImage-out.jpg';

        // Specify Cell Ranges and Save to certain Image formats
        sheet.ToImage(8, 1, 15, 5).Save(outputFileName);

        // Dispose of the workbook object to release resources
        workbook.Dispose();  
              
        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'image/jpeg'});

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
