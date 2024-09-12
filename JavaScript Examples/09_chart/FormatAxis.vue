<template>
  <span>Click the following button to set Chart Axis format </span>
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
       
        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "FormatAxis";

        _createChartData(sheet);

        //Add a chart
        let chart = sheet.Charts.Add({chartType: wasmModule.ExcelChartType.ColumnClustered});
        chart.DataRange = sheet.Range.get("B1:B9");
        chart.SeriesDataFromRange = false;
        chart.PlotArea.Visible = false;
        chart.TopRow = 10;
        chart.BottomRow = 28;
        chart.LeftColumn = 2;
        chart.RightColumn = 10;
        chart.ChartTitle = "Chart with Customized Axis";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;
        let cs1 = chart.Series.get(0);
        cs1.CategoryLabels = sheet.Range.get("A2:A9");

        //Format axis
        chart.PrimaryValueAxis.MajorUnit = 8;
        chart.PrimaryValueAxis.MinorUnit = 2;
        chart.PrimaryValueAxis.MaxValue = 50;
        chart.PrimaryValueAxis.MinValue = 0;
        chart.PrimaryValueAxis.IsReverseOrder = false;
        chart.PrimaryValueAxis.MajorTickMark = wasmModule.TickMarkType.TickMarkOutside;
        chart.PrimaryValueAxis.MinorTickMark = wasmModule.TickMarkType.TickMarkInside;
        chart.PrimaryValueAxis.TickLabelPosition = wasmModule.TickLabelPositionType.TickLabelPositionNextToAxis;
        chart.PrimaryValueAxis.CrossesAt = 0;

        //Set NumberFormat
        chart.PrimaryValueAxis.NumberFormat = "$#,##0";
        chart.PrimaryValueAxis.IsSourceLinked = false;

        let serie = chart.Series.get(0);
        for(let dataPoint of serie.DataPoints) {
            //Format Series
            dataPoint.DataFormat.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
            dataPoint.DataFormat.Fill.ForeColor = Module.wasmModule.Color.get_LightGreen();

            //Set transparency
            dataPoint.DataFormat.Fill.Transparency = 0.3;
        }

        // Save the modified workbook 
        const outputFile = 'FormatAxis.xlsx';
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
    function _createChartData (sheet) {
        //Set value of specified cell
        sheet.Range.get("A1").Value = "Month";
        sheet.Range.get("A2").Value = "Jan";
        sheet.Range.get("A3").Value = "Feb";
        sheet.Range.get("A4").Value = "Mar";
        sheet.Range.get("A5").Value = "Apr";
        sheet.Range.get("A6").Value = "May";
        sheet.Range.get("A7").Value = "Jun";
        sheet.Range.get("A8").Value = "Jul";
        sheet.Range.get("A9").Value = "Aug";

        sheet.Range.get("B1").Value = "Planned";
        sheet.Range.get("B2").NumberValue = 38;
        sheet.Range.get("B3").NumberValue = 47;
        sheet.Range.get("B4").NumberValue = 39;
        sheet.Range.get("B5").NumberValue = 36;
        sheet.Range.get("B6").NumberValue = 27;
        sheet.Range.get("B7").NumberValue = 25;
        sheet.Range.get("B8").NumberValue = 36;
        sheet.Range.get("B9").NumberValue = 48;
    }
    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
