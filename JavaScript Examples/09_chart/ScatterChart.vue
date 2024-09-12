<template>
  <span>Click the following button to create the Scatter Chart  </span>
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
        sheet.Name = "Scatter Chart";

        //Set chart data
        _CreateChartData(sheet);

        //Add a chart
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ScatterMarkers});

        //Set region of chart data
        chart.DataRange = sheet.Range.get("B2:B10");
        chart.SeriesDataFromRange = false;

        //Set position of chart
        chart.LeftColumn = 1;
        chart.TopRow = 11;
        chart.RightColumn = 10;
        chart.BottomRow = 28;

        chart.ChartTitle = "Scatter Chart";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;

        chart.Series.get(0).CategoryLabels = sheet.Range.get("A2:A10");
        chart.Series.get(0).Values = sheet.Range.get("B2:B10");

        //Add a trend line for the first series
        chart.Series.get(0).TrendLines.Add({type:wasmModule.TrendLineType.Exponential});

        chart.PrimaryValueAxis.Title = "Salary";
        chart.PrimaryCategoryAxis.Title = "Car Price";

        // Save the modified workbook 
        const outputFile = 'ScatterChart.xlsx';
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
      sheet.Range.get("A1").Value = "Y(Salary)";
      sheet.Range.get("A2").Value = "42763";
      sheet.Range.get("A3").Value = "195387";
      sheet.Range.get("A4").Value = "35672";
      sheet.Range.get("A5").Value = "217637";
      sheet.Range.get("A6").Value = "74734";
      sheet.Range.get("A7").Value = "130550";
      sheet.Range.get("A8").Value = "42976";
      sheet.Range.get("A9").Value = "15132";
      sheet.Range.get("A10").Value = "54936";

      sheet.Range.get("B1").Value = "X(Car Price)";
      sheet.Range.get("B2").Value = "19455";
      sheet.Range.get("B3").Value = "93965";
      sheet.Range.get("B4").Value = "20858";
      sheet.Range.get("B5").Value = "107164";
      sheet.Range.get("B6").Value = "34036";
      sheet.Range.get("B7").Value = "87806";
      sheet.Range.get("B8").Value = "17927";
      sheet.Range.get("B9").Value = "61518";
      sheet.Range.get("B10").Value = "29479";

      //Style
      sheet.Range.get("A1:B1").ColumnWidth = 12;
      sheet.Range.get("A1:B1").RowHeight = 15;
      sheet.Range.get("A1:B1").Style.Color = wasmModule.Color.get_DarkGray();
      sheet.Range.get("A1:B1").Style.Font.Color = wasmModule.Color.get_White();
      sheet.Range.get("A1:B1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Center;
      sheet.Range.get("A1:B1").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

      sheet.Range.get("A2:B10").Style.NumberFormat = "\"$\"#,##0";
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
