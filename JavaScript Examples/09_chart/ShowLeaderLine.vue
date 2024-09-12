<template>
  <span>Click the following button to make leader line of data label in chart show</span>
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
 
        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        // Set value of specified range
        sheet.Range.get("A1").Value = "1";
        sheet.Range.get("A2").Value = "2";
        sheet.Range.get("A3").Value = "3";
        sheet.Range.get("B1").Value = "4";
        sheet.Range.get("B2").Value = "5";
        sheet.Range.get("B3").Value = "6";
        sheet.Range.get("C1").Value = "7";
        sheet.Range.get("C2").Value = "8";
        sheet.Range.get("C3").Value = "9";

        let chart = sheet.Charts.Add({ chartType: wasmModule.ExcelChartType.BarStacked });
        chart.DataRange = sheet.Range.get("A1:C3");
        chart.TopRow = 4;
        chart.LeftColumn = 2;
        chart.Width = 450;
        chart.Height = 300;

        for (let cs of chart.Series) {
            cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
            cs.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;
        }

        // Save the modified workbook 
        const outputFile = 'ShowLeaderLine.xlsx';
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
