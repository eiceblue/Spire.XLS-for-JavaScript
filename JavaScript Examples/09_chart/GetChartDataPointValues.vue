<template>
  <span>Click the following button to get the values of chart data point in Excel file</span>
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
        let excelFileName = 'ChartToImage.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);
        const sb = [];
        //Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        //Get the chart
        let chart = sheet.Charts.get(0);

        //Get the first series of the chart
        let cs = chart.Series.get(0);

        for(let cr of cs.Values.Cells) {
            sb.push(cr.RangeAddress);

            //Get the data point value
            sb.push(`The value of the data point is ${cr.Value}`);
        }
       
        // Dispose of the workbook object to free resources
        workbook.Dispose();
        // // Create a Blob 
        const outputFile = 'GetChartDataPointValues.txt';
        const modifiedFile = new Blob([sb.toString()], { type: "text/plain;charset=utf-8"});

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
