<template>
  <span>Click the following button to create the Pareto Chart  </span>
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
        let excelFileName = 'ParetoChart.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);
        //Find the first worksheet
        let sheet = workbook.Worksheets.get(0);
        //Add chart
        let officeChart = sheet.Charts.Add();
        //Set chart type as Pareto
        officeChart.ChartType = wasmModule.ExcelChartType.Pareto;

        //Set data range in the worksheet
        officeChart.DataRange = sheet.Range.get("A2:B8");

        officeChart.TopRow = 1;
        officeChart.BottomRow = 19;
        officeChart.LeftColumn = 4;
        officeChart.RightColumn = 12;
        officeChart.PrimaryCategoryAxis.IsBinningByCategory = true;

        officeChart.PrimaryCategoryAxis.OverflowBinValue = 5;
        officeChart.PrimaryCategoryAxis.UnderflowBinValue = 1;

        //Formatting Pareto line
        officeChart.Series.get(0).ParetoLineFormat.LineProperties.Color = wasmModule.Color.get_Blue();

        //Gap width settings
        officeChart.Series.get(0).DataFormat.Options.GapWidth = 6;

        //Set the chart title
        officeChart.ChartTitle = "Expenses";

        // Save the modified workbook 
        const outputFile = 'CreateParetoChart.xlsx';
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
