<template>
  <span>Click the following button to get worksheet of a chart</span>
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

        //Access first worksheet of the workbook
        let worksheet = workbook.Worksheets.get(0);

        //Access the first chart inside this worksheet
        let chart = worksheet.Charts.get(0);

        //Get its worksheet
        let obj = chart.Worksheet;
        let wSheet = wasmModule.Worksheet.Convert(obj);

        //Create StringBuilder to save
        let content = [];

        //Set string format for displaying
        let result = `Sheet Name: ${worksheet.Name}\r\nCharts' sheet Name: ${wSheet.Name}`;

        //Add result string to StringBuilder
        content.push(result);
    
        // Dispose of the workbook object to free resources
        workbook.Dispose();
        // Create a Blob object 
        const outputFile = 'GetWorksheetOfChart.txt';
        const modifiedFile = new Blob([content.toString()],  { type: "text/plain;charset=utf-8"});

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
