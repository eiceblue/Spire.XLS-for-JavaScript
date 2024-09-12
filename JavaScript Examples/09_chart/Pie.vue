<template>
  <span>Click the following button to create pie chart </span>
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

        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "Pie Chart";

        //Add a chart
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Pie});

        //Set chart data
        _CreateChartData(sheet);

        //Set region of chart data
        chart.DataRange = sheet.Range.get("B2:B5");
        chart.SeriesDataFromRange = false;

        //Set position of chart
        chart.LeftColumn = 1;
        chart.TopRow = 6;
        chart.RightColumn = 9;
        chart.BottomRow = 25;

        //Chart title
        chart.ChartTitle = "Sales by year";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;

        let cs = chart.Series.get(0);
        cs.CategoryLabels = sheet.Range.get("A2:A5");
        cs.Values = sheet.Range.get("B2:B5");
        cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;

        chart.PlotArea.Fill.Visible = false;
       
        // Save the modified workbook 
        const outputFile = 'Pie.xlsx';
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
      sheet.Range.get("A1").Value = "Year";
      sheet.Range.get("A2").Value = "2002";
      sheet.Range.get("A3").Value = "2003";
      sheet.Range.get("A4").Value = "2004";
      sheet.Range.get("A5").Value = "2005";

      sheet.Range.get("B1").Value = "Sales";
      sheet.Range.get("B2").NumberValue = 4000;
      sheet.Range.get("B3").NumberValue = 6000;
      sheet.Range.get("B4").NumberValue = 7000;
      sheet.Range.get("B5").NumberValue = 8500;

      //Style
      sheet.Range.get("A1:B1").RowHeight = 15;
      sheet.Range.get("A1:B1").Style.Color = wasmModule.Color.get_DarkGray();
      sheet.Range.get("A1:B1").Style.Font.Color = wasmModule.Color.get_White();
      sheet.Range.get("A1:B1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Center;
      sheet.Range.get("A1:B1").Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

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
