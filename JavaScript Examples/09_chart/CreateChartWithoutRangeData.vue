<template>
  <span>Click the following button to create chart without range data</span>
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
        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        let sheet = workbook.Worksheets.get(0);

        // Add a chart to the worksheet
        let chart = sheet.Charts.Add();
        chart.ChartTitle = "Sample Chart";

        // Add a series to the chart
        let series = chart.Series.Add();

        // Add data
        series.EnteredDirectlyValues = [wasmModule.Int32.Create(10), wasmModule.Int32.Create(20), wasmModule.Int32.Create(30)];
        
        // Save the modified workbook 
        const outputFile = 'CreateChartWithoutRangeData.xlsx';
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
