<template>
  <span>Click the following button to add ClusteredBar chart</span>
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

        // Get the first sheet and set its name
        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "ClusteredBar";

        // Set chart data
        _CreateChartData(sheet);

        // Add a chart
        let chart = sheet.Charts.Add();

        // Set region of chart data
        chart.DataRange = sheet.Range.get("A1:C5");
        chart.SeriesDataFromRange = false;

        // Set position of chart
        chart.LeftColumn = 1;
        chart.TopRow = 6;
        chart.RightColumn = 11;
        chart.BottomRow = 29;
        chart.ChartType = wasmModule.ExcelChartType.BarClustered;

        // Chart title
        chart.ChartTitle = "Sales market by country";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;

        chart.PrimaryCategoryAxis.Title = "Country";
        chart.PrimaryCategoryAxis.Font.IsBold = true;
        chart.PrimaryCategoryAxis.TitleArea.IsBold = true;
        chart.PrimaryCategoryAxis.TitleArea.TextRotationAngle = 90;

        chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
        chart.PrimaryValueAxis.HasMajorGridLines = false;
        chart.PrimaryValueAxis.MinValue = 1000;
        chart.PrimaryValueAxis.TitleArea.IsBold = true;

        for (let cs of chart.Series) {
            cs.Format.Options.IsVaryColor = true;
            cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
        }

        chart.Legend.Position = wasmModule.LegendPositionType.Top;

        // Save the modified workbook 
        const outputFile = 'ClusteredBar.xlsx';
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
    function _CreateChartData(sheet) {
      // Set value of specified cell
      sheet.Range.get("A1").Value = "Country";
      sheet.Range.get("A2").Value = "Cuba";
      sheet.Range.get("A3").Value = "Mexico";
      sheet.Range.get("A4").Value = "France";
      sheet.Range.get("A5").Value = "German";

      sheet.Range.get("B1").Value = "Jun";
      sheet.Range.get("B2").NumberValue = 6000;
      sheet.Range.get("B3").NumberValue = 8000;
      sheet.Range.get("B4").NumberValue = 9000;
      sheet.Range.get("B5").NumberValue = 8500;

      sheet.Range.get("C1").Value = "Aug";
      sheet.Range.get("C2").NumberValue = 3000;
      sheet.Range.get("C3").NumberValue = 2000;
      sheet.Range.get("C4").NumberValue = 2300;
      sheet.Range.get("C5").NumberValue = 4200;

      // Style
      sheet.Range.get("A1:C1").RowHeight = 15;
      sheet.Range.get("A1:C1").Style.Color = wasmModule.Color.get_DarkGray();
      sheet.Range.get("A1:C1").Style.Font.Color = wasmModule.Color.get_White();
      sheet.Range.get("A1:C1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Center;
      sheet.Range.get("A1:C1").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

      sheet.Range.get("B2:C5").Style.NumberFormat = "\"$\"#,##0";
    }

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
