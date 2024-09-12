<template>
  <span>Click the following button to create radar chart in an excel workbook</span>
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
       
        //Initailize worksheet
        workbook.CreateEmptySheets(1);
        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "Chart data";
        sheet.GridLinesVisible = false;

        //Writes chart data
        CreateChartData(sheet);
        //Add a new  chart worsheet to workbook
        let chart = sheet.Charts.Add();

        //Set position of chart
        chart.LeftColumn = 1;
        chart.TopRow = 6;
        chart.RightColumn = 11;
        chart.BottomRow = 29;

        //Set region of chart data
        chart.DataRange = sheet.Range.get("A1:C5");
        chart.SeriesDataFromRange = false;

        chart.ChartType = wasmModule.ExcelChartType.Radar;

        //Chart title
        chart.ChartTitle = "Sale market by region";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;
        chart.PlotArea.Fill.Visible = false;
        chart.Legend.Position = wasmModule.LegendPositionType.Corner;
        
        // Save the modified workbook 
        const outputFile = 'CreateRadarChart.xlsx';
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
    function CreateChartData(sheet) {
        //Product
        sheet.Range.get("A1").Value = "Product";
        sheet.Range.get("A2").Value = "Bikes";
        sheet.Range.get("A3").Value = "Cars";
        sheet.Range.get("A4").Value = "Trucks";
        sheet.Range.get("A5").Value = "Buses";

        //Paris
        sheet.Range.get("B1").Value = "Paris";
        sheet.Range.get("B2").NumberValue = 4000;
        sheet.Range.get("B3").NumberValue = 23000;
        sheet.Range.get("B4").NumberValue = 4000;
        sheet.Range.get("B5").NumberValue = 30000;

        //New York
        sheet.Range.get("C1").Value = "New York";
        sheet.Range.get("C2").NumberValue = 30000;
        sheet.Range.get("C3").NumberValue = 7600;
        sheet.Range.get("C4").NumberValue = 18000;
        sheet.Range.get("C5").NumberValue = 8000;

        //Style
        sheet.Range.get("A1:C1").Style.Font.IsBold = true;
        sheet.Range.get("A2:C2").Style.KnownColor = wasmModule.ExcelColors.LightYellow;
        sheet.Range.get("A3:C3").Style.KnownColor = wasmModule.ExcelColors.LightGreen1;
        sheet.Range.get("A4:C4").Style.KnownColor = wasmModule.ExcelColors.LightOrange;
        sheet.Range.get("A5:C5").Style.KnownColor = wasmModule.ExcelColors.LightTurquoise;

        //Border
        let style = sheet.Range.get("A1:C5").Style;
        let borders = style.Borders;
        let topborder = borders.get(wasmModule.BordersLineType.EdgeTop);
        topborder.Color = wasmModule.Color.FromArgb(0, 0, 128);
        borders.get(wasmModule.BordersLineType.EdgeTop).LineStyle = wasmModule.LineStyleType.Thin;
        borders.get(wasmModule.BordersLineType.EdgeBottom).Color = wasmModule.Color.FromArgb(0, 0, 128);
        sheet.Range.get("A1:C5").Style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle = wasmModule.LineStyleType.Thin;
        sheet.Range.get("A1:C5").Style.Borders.get(wasmModule.BordersLineType.EdgeLeft).Color = wasmModule.Color.FromArgb(0, 0, 128);
        sheet.Range.get("A1:C5").Style.Borders.get(wasmModule.BordersLineType.EdgeLeft).LineStyle = wasmModule.LineStyleType.Thin;
        sheet.Range.get("A1:C5").Style.Borders.get(wasmModule.BordersLineType.EdgeRight).Color = wasmModule.Color.FromArgb(0, 0, 128);
        sheet.Range.get("A1:C5").Style.Borders.get(wasmModule.BordersLineType.EdgeRight).LineStyle = wasmModule.LineStyleType.Thin;

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
