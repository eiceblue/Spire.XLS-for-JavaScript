<template>
  <span>Click the following button to create multi level category chart</span>
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

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        let sheet = workbook.Worksheets.get(0);

        //Write data to cells
        sheet.Range.get("A1").Text = "Main Category";
        sheet.Range.get("A2").Text = "Fruit";
        sheet.Range.get("A6").Text = "Vegies";
        sheet.Range.get("B1").Text = "Sub Category";
        sheet.Range.get("B2").Text = "Bananas";
        sheet.Range.get("B3").Text = "Oranges";
        sheet.Range.get("B4").Text = "Pears";
        sheet.Range.get("B5").Text = "Grapes";
        sheet.Range.get("B6").Text = "Carrots";
        sheet.Range.get("B7").Text = "Potatoes";
        sheet.Range.get("B8").Text = "Celery";
        sheet.Range.get("B9").Text = "Onions";
        sheet.Range.get("C1").Text = "Value";
        sheet.Range.get("C2").Value = "52";
        sheet.Range.get("C3").Value = "65";
        sheet.Range.get("C4").Value = "50";
        sheet.Range.get("C5").Value = "45";
        sheet.Range.get("C6").Value = "64";
        sheet.Range.get("C7").Value = "62";
        sheet.Range.get("C8").Value = "89";
        sheet.Range.get("C9").Value = "57";

        //Vertically merge cells from A2 to A5, A6 to A9
        sheet.Range.get("A2:A5").Merge();
        sheet.Range.get("A6:A9").Merge();
        sheet.AutoFitColumn(1);
        sheet.AutoFitColumn(2);

        //Add a clustered bar chart to worksheet
        let chart = sheet.Charts.Add({chartType: wasmModule.ExcelChartType.BarClustered});
        chart.ChartTitle = "Value";
        chart.PlotArea.Fill.FillType = wasmModule.ShapeFillType.NoFill;
        chart.Legend.Delete();
        chart.LeftColumn = 5;
        chart.TopRow = 1;
        chart.RightColumn = 14;

        //Set the data source of series data
        chart.DataRange = sheet.Range.get("C2:C9");
        chart.SeriesDataFromRange = false;
        //Set the data source of category labels
        let serie = chart.Series.get(0);
        serie.CategoryLabels = sheet.Range.get("A2:B9");
        //Show multi-level category labels
        chart.PrimaryCategoryAxis.MultiLevelLable = true;
       
        // Save the modified workbook 
        const outputFile = 'CreateMultiLevelChart.xlsx';
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
