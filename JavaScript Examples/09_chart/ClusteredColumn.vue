<template>
  <span>Click the following button to ClusteredColumn chart</span>
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
        const sheet = workbook.Worksheets.get(0);
        sheet.Name = "ClusteredColumn";

        _CreateChartData(sheet);

        const chart = sheet.Charts.Add();
        chart.DataRange = sheet.Range.get("A1:C5");
        chart.SeriesDataFromRange = false;

        chart.LeftColumn = 1;
        chart.TopRow = 6;
        chart.RightColumn = 11;
        chart.BottomRow = 29;

        chart.ChartType = wasmModule.ExcelChartType.ColumnClustered;

        chart.ChartTitle = "Sales market by country";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;

        chart.PrimaryCategoryAxis.Title = "Country";
        chart.PrimaryCategoryAxis.Font.IsBold = true;
        chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

        chart.PrimaryValueAxis.Title = "Sales(in Dollars)";
        chart.PrimaryValueAxis.HasMajorGridLines = false;
        chart.PrimaryValueAxis.MinValue = 1000;
        chart.PrimaryValueAxis.TitleArea.IsBold = true;
        chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;

        for (let i = 0; i < chart.Series.Length; i++) {
            let cs = chart.Series.get(i);
            cs.Format.Options.IsVaryColor = true;
            cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
        }

        chart.Legend.Position = wasmModule.LegendPositionType.Top;

        // Save the modified workbook 
        const outputFile = 'ClusteredColumn.xlsx';
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
