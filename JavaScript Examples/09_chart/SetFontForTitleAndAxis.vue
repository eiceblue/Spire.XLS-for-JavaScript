<template>
  <span>Click the following button to set font for chart title and chart axis</span>
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

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Set font for chart title and chart axis
        let worksheet = workbook.Worksheets.get(0);
        let chart = worksheet.Charts.get(0);

        // Format the font for the chart title
        chart.ChartTitleArea.Color = wasmModule.Color.get_Blue();
        chart.ChartTitleArea.Size = 20.0;

        // Format the font for the chart Axis
        chart.PrimaryValueAxis.Font.Color = wasmModule.Color.get_Gold();
        chart.PrimaryValueAxis.Font.Size = 10.0;
        chart.PrimaryCategoryAxis.Font.Color = wasmModule.Color.get_Red();
        chart.PrimaryCategoryAxis.Font.Size = 20.0;
        
        // Save the modified workbook 
        const outputFile = 'SetFontForTitleAndAxis.xlsx';
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
