<template>
  <span>Click the following button to create pivot chart</span>
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
        // Fetch the font file and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);
        // Fetch the Excel file and add it to the Virtual File System (VFS)
        let excelFileName = 'PivotTable.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //get the first worksheet
        let sheet = workbook.Worksheets.get(0);
        //get the first pivot table in the worksheet
        let pivotTable = sheet.PivotTables.get(0);

        //create a clustered column chart based on the pivot table
        let chart = sheet.Charts.Add({pivotChartType: wasmModule.ExcelChartType.ColumnClustered, pivotTable: pivotTable});
        //set chart position
        chart.TopRow = 10;
        chart.LeftColumn = 1;
        chart.RightColumn = 7;
        chart.BottomRow = 25;
        //set chart title
        chart.ChartTitle = "Pivot Chart";
        
        // Save the modified workbook 
        const outputFile = 'CreatePivotChart.xlsx';
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
