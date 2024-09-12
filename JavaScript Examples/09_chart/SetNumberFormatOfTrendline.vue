<template>
  <span>Click the following button to set number format of trendline</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref, onMounted, onUnmounted } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");
    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Fetch the Excel file and add it to the Virtual File System (VFS)
        let excelFileName = 'ChartSample4.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        

        workbook.LoadFromFile(excelFileName);

        // Get the chart from the first worksheet
        let chart = workbook.Worksheets.get(0).Charts.get(0);

        // Get the trendline of the chart and then extract the equation of the trendline
        let trendLine = chart.Series.get(1).TrendLines.get(0);

        // Set the number format of trendLine to "#,##0.00"
        trendLine.DataLabel.NumberFormat = "#,##0.00";

        // Save the modified workbook 
        const outputFile = 'SetNumberFormatOfTrendline.xlsx';
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
