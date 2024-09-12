<template>
  <span>Click the following button to create a doughnut chart</span>
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

        // Insert data
        sheet.Range.get("A1").Value = "Country";
        sheet.Range.get("A1").Style.Font.IsBold = true;
        sheet.Range.get("A2").Value = "Cuba";
        sheet.Range.get("A3").Value = "Mexico";
        sheet.Range.get("A4").Value = "France";
        sheet.Range.get("A5").Value = "German";
        sheet.Range.get("B1").Value = "Sales";
        sheet.Range.get("B1").Style.Font.IsBold = true;
        sheet.Range.get("B2").NumberValue = 6000;
        sheet.Range.get("B3").NumberValue = 8000;
        sheet.Range.get("B4").NumberValue = 9000;
        sheet.Range.get("B5").NumberValue = 8500;

        // Add a new chart, set chart type as doughnut
        let chart = sheet.Charts.Add();
        chart.ChartType = wasmModule.ExcelChartType.Doughnut;
        chart.DataRange = sheet.Range.get("A1:B5");
        chart.SeriesDataFromRange = false;

        // Set position of chart
        chart.LeftColumn = 4;
        chart.TopRow = 2;
        chart.RightColumn = 12;
        chart.BottomRow = 22;

        // Chart title
        chart.ChartTitle = "Market share by country";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;

        for (let i = 0; i < chart.Series.Count; i++) {
            chart.Series.get(i).DataPoints.DefaultDataPoint.DataLabels.HasPercentage = true;
        }

        chart.Legend.Position = wasmModule.LegendPositionType.Top;

        // Save the modified workbook
        const outputFile = 'CreateDoughnutChart.xlsx';
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
