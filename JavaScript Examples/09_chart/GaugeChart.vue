<template>
  <span>Click the following button to create Gauge Chart</span>
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

        //Get the first sheet and set its name
        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "Gauge Chart";

        //Set chart data
        _CreateChartData(sheet);

        //Add a Doughnut chart
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Doughnut});
        chart.DataRange = sheet.Range.get("A1:A5");
        chart.SeriesDataFromRange = false;
        chart.HasLegend = true;

        //Set the position of chart
        chart.LeftColumn = 2;
        chart.TopRow = 7;
        chart.RightColumn = 9;
        chart.BottomRow = 25;

        //Get the series 1
        let cs1 = chart.Series.get({name:"Value"});
        cs1.Format.Options.DoughnutHoleSize = 60;
        cs1.DataFormat.Options.FirstSliceAngle = 270;

        //Set the fill color
        cs1.DataPoints.get(0).DataFormat.Fill.ForeColor = wasmModule.Color.get_Yellow();
        cs1.DataPoints.get(1).DataFormat.Fill.ForeColor = wasmModule.Color.get_PaleVioletRed
        cs1.DataPoints.get(2).DataFormat.Fill.ForeColor = wasmModule.Color.get_DarkViolet();
        cs1.DataPoints.get(3).DataFormat.Fill.Visible = false;

        //Add a series with pie chart
        let cs2 = chart.Series.Add({name:"Pointer", serieType:wasmModule.ExcelChartType.Pie});

        //Set the value
        cs2.Values = sheet.Range.get("D2:D4");
        cs2.UsePrimaryAxis = false;
        cs2.DataPoints.get(0).DataLabels.HasValue = true;
        cs2.DataFormat.Options.FirstSliceAngle = 270;
        cs2.DataPoints.get(0).DataFormat.Fill.Visible = false;
        cs2.DataPoints.get(1).DataFormat.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
        cs2.DataPoints.get(1).DataFormat.Fill.ForeColor = wasmModule.Color.get_Black();
        cs2.DataPoints.get(2).DataFormat.Fill.Visible = false;

        // Save the modified workbook 
        const outputFile = 'GaugeChart.xlsx';
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
    //Set value of specified cell
    sheet.Range.get("A1").Value = "Value";
    sheet.Range.get("A2").Value = "30";
    sheet.Range.get("A3").Value = "60";
    sheet.Range.get("A4").Value = "90";
    sheet.Range.get("A5").Value = "180";
    sheet.Range.get("C2").Value = "value";
    sheet.Range.get("C3").Value = "pointer";
    sheet.Range.get("C4").Value = "End";
    sheet.Range.get("D2").Value = "10";
    sheet.Range.get("D3").Value = "1";
    sheet.Range.get("D4").Value = "189";
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
