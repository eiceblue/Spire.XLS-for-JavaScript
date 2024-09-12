<template>
  <span>Click the following button to set color for chart area</span>
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
     
        let excelFileName = 'SampeB_4.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Get the chart
        let chart = sheet.Charts.get(0);

        // Set color for chart area
        chart.ChartArea.Fill.ForeColor = wasmModule.Color.get_LightSeaGreen();

        // Set color for plot area
        chart.PlotArea.Fill.ForeColor = wasmModule.Color.get_LightGray();

        // Save the modified workbook 
        const outputFile = 'SetColorForChartArea.xlsx';
        // Dispose of the workbook object to free resources
        workbook.SaveToFile(outputFile);
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
