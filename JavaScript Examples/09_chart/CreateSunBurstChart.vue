<template>
  <span>Click the following button to create the SunBurst Chart.</span>
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
        let excelFileName = 'SunBurst.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);
        //Find the first worksheet
        let sheet = workbook.Worksheets.get(0);
        //Add chart
        let officeChart = sheet.Charts.Add();
        //Set chart type as Sunburst
        officeChart.ChartType = wasmModule.ExcelChartType.SunBurst;

        //Set data range in the worksheet
        officeChart.DataRange = sheet.Range.get("A1:D16");

        officeChart.TopRow = 1;
        officeChart.BottomRow = 17;
        officeChart.LeftColumn = 6;
        officeChart.RightColumn = 14;

        //Set the chart title
        officeChart.ChartTitle = "Sales by quarter";

        //Formatting data labels
        officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;

        //Hiding the legend
        officeChart.HasLegend = false;

        // Save the modified workbook 
        const outputFile = 'CreateSunBurstChart.xlsx';
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
