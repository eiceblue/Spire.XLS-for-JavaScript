<template>
  <span>Click the following button to Add Error Bars for Chart</span>
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
        //Create a empty sheet
        workbook.CreateEmptySheets(1);
        //Add data
        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "Demo";
        sheet.Range.get("A1").Value = "Month";
        sheet.Range.get("A2").Value = "Jan.";
        sheet.Range.get("A3").Value = "Feb.";
        sheet.Range.get("A4").Value = "Mar.";
        sheet.Range.get("A5").Value = "Apr.";
        sheet.Range.get("A6").Value = "May.";
        sheet.Range.get("A7").Value = "Jun.";
        sheet.Range.get("B1").Value = "Planned";
        sheet.Range.get("B2").NumberValue = 3.3;
        sheet.Range.get("B3").NumberValue = 2.5;
        sheet.Range.get("B4").NumberValue = 2.0;
        sheet.Range.get("B5").NumberValue = 3.7;
        sheet.Range.get("B6").NumberValue = 4.5;
        sheet.Range.get("B7").NumberValue = 4.0;
        sheet.Range.get("C1").Value = "Actual";
        sheet.Range.get("C2").NumberValue = 3.8;
        sheet.Range.get("C3").NumberValue = 3.2;
        sheet.Range.get("C4").NumberValue = 1.7;
        sheet.Range.get("C5").NumberValue = 3.5;
        sheet.Range.get("C6").NumberValue = 4.5;
        sheet.Range.get("C7").NumberValue = 4.3;

        //Add a line chart and then add percentage error bar to the chart
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.Line});
        chart.DataRange = sheet.Range.get("B1:B7");
        chart.SeriesDataFromRange = false;
        //Set chart position
        chart.TopRow = 8;
        chart.BottomRow = 25;
        chart.LeftColumn = 2;
        chart.RightColumn = 9;
        chart.ChartTitle = "Error Bar 10% Plus";
        chart.ChartTitleArea.IsBold = true;
        chart.ChartTitleArea.Size = 12;
        let cs1 = chart.Series.get(0);
        cs1.CategoryLabels = sheet.Range.get("A2:A7");
        cs1.ErrorBar({bIsY:true, include:wasmModule.ErrorBarIncludeType.Plus, type:wasmModule.ErrorBarType.Percentage, numberValue:10.0});

        // Add a column chart with standard error bars as comparison
        let chart2 = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ColumnClustered});
        chart2.DataRange = sheet.Range.get("B1:C7");
        chart2.SeriesDataFromRange = false;

        //Set chart position
        chart2.TopRow = 8;
        chart2.BottomRow = 25;
        chart2.LeftColumn = 10;
        chart2.RightColumn = 17;
        chart2.ChartTitle = "Standard Error Bar";
        chart2.ChartTitleArea.IsBold = true;
        chart2.ChartTitleArea.Size = 12;
        let cs2 = chart2.Series.get(0);
        cs2.CategoryLabels = sheet.Range.get("A2:A7");
        cs2.ErrorBar({bIsY:true, include:wasmModule.ErrorBarIncludeType.Minus, type:wasmModule.ErrorBarType.StandardError, numberValue:0.3});
        let cs3 = chart2.Series.get(1);
        cs3.ErrorBar({bIsY:true, include:wasmModule.ErrorBarIncludeType.Both, type:wasmModule.ErrorBarType.StandardError, numberValue:0.5});
 
        // Save the modified workbook
        const outputFile = 'AddErrorBars.xlsx';
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
