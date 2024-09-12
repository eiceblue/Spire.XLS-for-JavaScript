<template>
  <span>Click the following button to set customized data marker for charts</span>
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
        //Add some sample data
        sheet.Name = "Demo";
        sheet.Range.get("A1").Value = "Tom";
        sheet.Range.get("A2").NumberValue = 1.5;
        sheet.Range.get("A3").NumberValue = 2.1;
        sheet.Range.get("A4").NumberValue = 3.6;
        sheet.Range.get("A5").NumberValue = 5.2;
        sheet.Range.get("A6").NumberValue = 7.3;
        sheet.Range.get("A7").NumberValue = 3.1;
        sheet.Range.get("B1").Value = "Kitty";
        sheet.Range.get("B2").NumberValue = 2.5;
        sheet.Range.get("B3").NumberValue = 4.2;
        sheet.Range.get("B4").NumberValue = 1.3;
        sheet.Range.get("B5").NumberValue = 3.2;
        sheet.Range.get("B6").NumberValue = 6.2;
        sheet.Range.get("B7").NumberValue = 4.7;

        //Create a Scatter-Markers chart based on the sample data
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ScatterMarkers});
        chart.DataRange = sheet.Range.get("A1:B7");
        chart.PlotArea.Visible = false;
        chart.SeriesDataFromRange = false;
        chart.TopRow = 5;
        chart.BottomRow = 22;
        chart.LeftColumn = 4;
        chart.RightColumn = 11;
        chart.ChartTitle = "Chart with Markers";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 10;

        //Format the markers in the chart by setting the background color, foreground color, type, size and transparency
        let cs1 = chart.Series.get(0);
        cs1.DataFormat.MarkerBackgroundColor = wasmModule.Color.get_RoyalBlue();
        cs1.DataFormat.MarkerForegroundColor = wasmModule.Color.get_WhiteSmoke();
        cs1.DataFormat.MarkerSize = 7;
        cs1.DataFormat.MarkerStyle = wasmModule.ChartMarkerType.PlusSign;
        cs1.DataFormat.MarkerTransparencyValue = 0.8;

        let cs2 = chart.Series.get(1);
        cs2.DataFormat.MarkerBackgroundColor = wasmModule.Color.get_Pink();
        cs2.DataFormat.MarkerSize = 9;
        cs2.DataFormat.MarkerStyle = wasmModule.ChartMarkerType.Triangle;
        cs2.DataFormat.MarkerTransparencyValue = 0.9;

        // Save the modified workbook 
        const outputFile = 'CustomDataMarker.xlsx';
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
