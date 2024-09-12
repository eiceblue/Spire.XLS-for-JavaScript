<template>
  <span>Click the following button to create bubble chart</span>
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
        let excelFileName = 'CreateBubbleChart.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first sheet and set its name
        let sheet = workbook.Worksheets.get(0);

        // Add a chart
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Bubble});

        // Set region of chart data
        chart.DataRange = sheet.Range.get("A1:C5");
        chart.SeriesDataFromRange = false;
        chart.Series.get(0).Bubbles = sheet.Range.get("C2:C5");

        // Set position of chart
        chart.LeftColumn = 7;
        chart.TopRow = 6;
        chart.RightColumn = 16;
        chart.BottomRow = 29;

        chart.ChartTitle = "Bubble Chart";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;

        // Save the modified workbook 
        const outputFile = 'CreateBubbleChart.xlsx';
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
