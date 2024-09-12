<template>
  <span>Click the following button to create the Funnel Chart</span>
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
        let excelFileName = 'Funnel.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);
        //Find the first worksheet
        let sheet = workbook.Worksheets.get(0);
        //Add a new chart
        let officeChart = sheet.Charts.Add();
        //Set chart type as Funnel
        officeChart.ChartType = wasmModule.ExcelChartType.Funnel;

        //Set data range in the worksheet
        officeChart.DataRange = sheet.Range.get("A1:B6");

        //Set the chart title
        officeChart.ChartTitle = "Funnel";

        //Formatting the legend and data label option
        officeChart.HasLegend = false;
        officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.HasValue = true;
        officeChart.Series.get(0).DataPoints.DefaultDataPoint.DataLabels.Size = 8;

        // Save the modified workbook 
        const outputFile = 'CreateFunnelChart.xlsx';
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
