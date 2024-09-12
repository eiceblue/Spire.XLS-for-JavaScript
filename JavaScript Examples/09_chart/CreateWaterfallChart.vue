<template>
  <span>Click the following button to create the Waterfall Chart</span>
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
        let excelFileName = 'WaterfallChart.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 

        workbook.LoadFromFile(excelFileName);
        let sheet = workbook.Worksheets.get(0);
        let officeChart = sheet.Charts.Add();
        //Set chart type as waterfall
        officeChart.ChartType = wasmModule.ExcelChartType.WaterFall;

        //Set data range to the chart from the worksheet
        officeChart.DataRange = sheet.Range.get("A2:B8");

        officeChart.TopRow = 1;
        officeChart.BottomRow = 19;
        officeChart.LeftColumn = 4;
        officeChart.RightColumn = 12;

        //Data point settings as total in chart
        officeChart.Series.get(0).DataPoints.get(3).SetAsTotal = true;
        officeChart.Series.get(0).DataPoints.get(6).SetAsTotal = true;

        //Showing the connector lines between data points
        officeChart.Series.get(0).Format.ShowConnectorLines = true;

        //Set the chart title
        officeChart.ChartTitle = "WaterFall Chart";

        //Formatting data label and legend option
        officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
        officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;
        officeChart.Legend.Position = wasmModule.LegendPositionType.Right;

        // Save the modified workbook 
        const outputFile = 'CreateWaterfallChart.xlsx';
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
