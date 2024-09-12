<template>
  <span>Click the following button to set discontinuous data for chart serie</span>
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
        let excelFileName = 'DiscontinuousData.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        //Add a chart
        let chart = sheet.Charts.Add({chartType:wasmModule.ExcelChartType.ColumnClustered});
        chart.SeriesDataFromRange = false;

        //Set the position of chart
        chart.LeftColumn = 1;
        chart.TopRow = 10;
        chart.RightColumn = 10;
        chart.BottomRow = 24;

        //Add a series
        let cs1 = chart.Series.Add();

        //Set the name of the cs1
        cs1.Name = sheet.Range.get("B1").Value;

        //Set discontinuous values for cs1
        cs1.CategoryLabels = sheet.Range.get("A2:A3").AddCombinedRange(sheet.Range.get("A5:A6")).AddCombinedRange(sheet.Range.get("A8:A9"));
        cs1.Values = sheet.Range.get("B2:B3").AddCombinedRange(sheet.Range.get("B5:B6")).AddCombinedRange(sheet.Range.get("B8:B9"));

        //Set the chart type
        cs1.SerieType = wasmModule.ExcelChartType.ColumnClustered;

        //Add a series
        let cs2 = chart.Series.Add();
        cs2.Name = sheet.Range.get("C1").Value;
        cs2.CategoryLabels = sheet.Range.get("A2:A3").AddCombinedRange(sheet.Range.get("A5:A6")).AddCombinedRange(sheet.Range.get("A8:A9"));
        cs2.Values = sheet.Range.get("C2:C3").AddCombinedRange(sheet.Range.get("C5:C6")).AddCombinedRange(sheet.Range.get("C8:C9"));
        cs2.SerieType = wasmModule.ExcelChartType.ColumnClustered;

        chart.ChartTitle = "Chart";
        chart.ChartTitleArea.Font.Size = 20;
        chart.ChartTitleArea.Color = wasmModule.Color.get_Black();

        chart.PrimaryValueAxis.HasMajorGridLines = false;

        // Save the modified workbook 
        const outputFile = 'DiscontinuousData.xlsx';
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
