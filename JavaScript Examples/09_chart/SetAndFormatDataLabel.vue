<template>
  <span>Click the following button to set and format data labels for charts</span>
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
        workbook.CreateEmptySheets(1);
        let sheet = workbook.Worksheets.get(0);

        sheet.Name = "Demo";
        sheet.Range.get("A1").Value = "Month";
        sheet.Range.get("A2").Value = "Jan";
        sheet.Range.get("A3").Value = "Feb";
        sheet.Range.get("A4").Value = "Mar";
        sheet.Range.get("A5").Value = "Apr";
        sheet.Range.get("A6").Value = "May";
        sheet.Range.get("A7").Value = "Jun";
        sheet.Range.get("B1").Value = "Peter";
        sheet.Range.get("B2").NumberValue = 25;
        sheet.Range.get("B3").NumberValue = 18;
        sheet.Range.get("B4").NumberValue = 8;
        sheet.Range.get("B5").NumberValue = 13;
        sheet.Range.get("B6").NumberValue = 22;
        sheet.Range.get("B7").NumberValue = 28;

        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.LineMarkers});
        chart.DataRange = sheet.Range.get("B1:B7");
        chart.PlotArea.Visible = false;
        chart.SeriesDataFromRange = false;
        chart.TopRow = 5;
        chart.BottomRow = 26;
        chart.LeftColumn = 2;
        chart.RightColumn = 11;
        chart.ChartTitle = "Data Labels Demo";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;
        let cs1 = chart.Series.get(0);
        cs1.CategoryLabels = sheet.Range.get("A2:A7");

        cs1.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
        cs1.DataPoints.DefaultDataPoint.DataLabels.HasLegendKey = false;
        cs1.DataPoints.DefaultDataPoint.DataLabels.HasPercentage = false;
        cs1.DataPoints.DefaultDataPoint.DataLabels.HasSeriesName = true;
        cs1.DataPoints.DefaultDataPoint.DataLabels.HasCategoryName = true;
        cs1.DataPoints.DefaultDataPoint.DataLabels.Delimiter = ". ";

        cs1.DataPoints.DefaultDataPoint.DataLabels.Size = 9;
        cs1.DataPoints.DefaultDataPoint.DataLabels.Color = wasmModule.Color.get_Red();
        cs1.DataPoints.DefaultDataPoint.DataLabels.FontName = "Calibri";
        cs1.DataPoints.DefaultDataPoint.DataLabels.Position = wasmModule.DataLabelPositionType.Center;

        // Save the modified workbook 
        const outputFile = 'SetAndFormatDataLabel.xlsx';
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
