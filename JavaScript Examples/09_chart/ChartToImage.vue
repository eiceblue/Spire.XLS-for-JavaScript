<template>
  <span>Click the following button to save chart as image </span>
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
        // Fetch the font file and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);
        // Fetch the Excel file and add it to the Virtual File System (VFS)
        let excelFileName = 'ChartToImage.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        
        workbook.LoadFromFile(excelFileName);

        //Save chart as image
        const image = workbook.SaveChartAsImage({worksheet:workbook.Worksheets.get(0), chartIndex:0});
        const outputFile = 'ChartToImage.png';
        image.Save(outputFile);

        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved image from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/png' });

        // Download the image
        downloadName.value = outputFile;
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
