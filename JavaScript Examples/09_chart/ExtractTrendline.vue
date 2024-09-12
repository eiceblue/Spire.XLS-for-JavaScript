<template>
  <span>Click the following button to extract the trendline equation from chart</span>
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
        let excelFileName = 'ChartSample4.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the chart from the first worksheet
        let chart = workbook.Worksheets.get(0).Charts.get(0);

        //Get the trendline of the chart and then extract the equation of the trendline
        let trendLine = chart.Series.get(1).TrendLines.get(0);
        let formula = trendLine.Formula;
        let equation = "The equation is: " + formula;

        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Create a Blob object with the equation string
        const outputFile = 'ExtractTrendline.txt';
        const modifiedFile = new Blob([equation], { type: "text/plain;charset=utf-8" });

        // Download the txt file
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
