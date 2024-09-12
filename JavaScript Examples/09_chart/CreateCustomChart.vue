<template>
  <span>Click the following button to create custom chart</span>
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

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Set values
        sheet.Range.get("A1").Value = "60";
        sheet.Range.get("A2").Value = "90";
        sheet.Range.get("A3").Value = "80";
        sheet.Range.get("A4").Value = "85";
        sheet.Range.get("B1").Value = "100";
        sheet.Range.get("B2").Value = "110";
        sheet.Range.get("B3").Value = "80";
        sheet.Range.get("B4").Value = "70";

        // Add a chart based on the data from A1 to B4
        let chart = sheet.Charts.Add();
        chart.DataRange = sheet.Range.get("A1:B4");
        chart.SeriesDataFromRange = false;

        // Set position of chart
        chart.LeftColumn = 1;
        chart.TopRow = 10;
        chart.RightColumn = 7;
        chart.BottomRow = 25;

        // Apply different chart type to different series
        let cs1 = chart.Series.get(0);
        cs1.SerieType = wasmModule.ExcelChartType.ColumnClustered;
        let cs2 = chart.Series.get(1);
        cs2.SerieType = wasmModule.ExcelChartType.Line;

        chart.ChartTitle = "Custom chart";
        
        // Save the modified workbook         
        const outputFile = 'CreateCustomChart.xlsx';
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
