<template>
  <span>Click the following button to create the BoxAndWhisker Chart</span>
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
        // Fetch the Excel file and add it to the Virtual File System (VFS)     
        let excelFileName = 'BoxAndWhiskerChart.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);
        let sheet = workbook.Worksheets.get(0);

        // Add a new chart
        let officeChart = sheet.Charts.Add();

        // Set the chart title
        officeChart.ChartTitle = "Yearly Vehicle Sales";

        // Set chart type as Box and Whisker
        officeChart.ChartType = wasmModule.ExcelChartType.BoxAndWhisker;

        // Set data range in the worksheet
        officeChart.DataRange = sheet.Range.get("A1:E17");

        // Box and Whisker settings on first series
        let seriesA = officeChart.Series.get(0);
        seriesA.DataFormat.ShowInnerPoints = false;
        seriesA.DataFormat.ShowOutlierPoints = true;
        seriesA.DataFormat.ShowMeanMarkers = true;
        seriesA.DataFormat.ShowMeanLine = false;
        seriesA.DataFormat.QuartileCalculationType = wasmModule.ExcelQuartileCalculation.ExclusiveMedian;

        // Box and Whisker settings on second series
        let seriesB = officeChart.Series.get(1);
        seriesB.DataFormat.ShowInnerPoints = false;
        seriesB.DataFormat.ShowOutlierPoints = true;
        seriesB.DataFormat.ShowMeanMarkers = true;
        seriesB.DataFormat.ShowMeanLine = false;
        seriesB.DataFormat.QuartileCalculationType = wasmModule.ExcelQuartileCalculation.InclusiveMedian;

        // Box and Whisker settings on third series
        let seriesC = officeChart.Series.get(2);
        seriesC.DataFormat.ShowInnerPoints = false;
        seriesC.DataFormat.ShowOutlierPoints = true;
        seriesC.DataFormat.ShowMeanMarkers = true;
        seriesC.DataFormat.ShowMeanLine = false;
        seriesC.DataFormat.QuartileCalculationType = wasmModule.ExcelQuartileCalculation.ExclusiveMedian;

        // Save the modified workbook 
        const outputFile = 'BoxAndWhiskerChart.xlsx';
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
