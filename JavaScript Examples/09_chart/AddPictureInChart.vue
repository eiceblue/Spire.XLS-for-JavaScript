<template>
  <span>Click the following button to add a picture in the chart </span>
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
        // Fetch the image and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS('SpireXls.png', '', `${import.meta.env.BASE_URL}static/image/`);
        // Fetch the Excel file and add it to the Virtual File System (VFS)
        let excelFileName = 'ChartToImage.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        //Get the chart
        let chart = sheet.Charts.get(0);

        //Add the picture in chart
        chart.Shapes.AddPicture("SpireXls.png");

        // Save the modified workbook
        const outputFile = 'AddPictureInChart.xlsx';
        workbook.SaveToFile(outputFile);
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved Excel file from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Download the Excel file
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
