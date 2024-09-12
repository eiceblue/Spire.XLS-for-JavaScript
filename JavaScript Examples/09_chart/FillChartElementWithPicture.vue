<template>
  <span>Click the following button to fill chart elements with picture</span>
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
        // Fetch the Excel file and add it to the Virtual File System (VFS)
        let excelFileName = 'ChartSample1.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        // Fetch the image and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS("Background.png", '', `${import.meta.env.BASE_URL}static/image/`);
        
        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
     
        workbook.LoadFromFile(excelFileName);

        let ws = workbook.Worksheets.get(0);
        let chart = ws.Charts.get(0);

        // Fill plot area with image
        // chart.ChartArea.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile("Background.png"), name:"None"});

        chart.PlotArea.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile("Background.png"), name:"None"});

        // Save the modified workbook 
        const outputFile = 'FillChartElementWithPicture.xlsx';
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
