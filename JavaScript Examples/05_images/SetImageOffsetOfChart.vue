<template>
  <span>Click the following button to set the offset of image when the filling way of chart is picture filing</span>
    <el-button @click="startProcessing">Start</el-button>
    <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
      Click here to download the generated file
    </a>
  </template>
<script>
import { ref} from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      if (wasmModule) {
        wasmModule = window.wasmModule;
        let inputFileName1='SetImageOffsetOfChart.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName1, '', `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2='Background.png';
        await wasmModule.FetchFileToVFS(inputFileName2, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName1});
           
        // Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);
        let sheet1 = workbook.Worksheets.Add("Contrast");

        // Add chart1 and background image to sheet1 as comparison.
        let chart1 = sheet1.Charts.Add({chartType:wasmModule.ExcelChartType.ColumnClustered});
        chart1.DataRange = sheet.Range.get("D1:E8");
        chart1.SeriesDataFromRange = false;

        // Chart Position.
        chart1.LeftColumn = 1;
        chart1.TopRow = 11;
        chart1.RightColumn = 8;
        chart1.BottomRow = 33;

        let bm = wasmModule.Stream.CreateByFile(inputFileName2);
        // Add picture as background.
        chart1.ChartArea.Fill.CustomPicture({im:bm,name:"None"});
        chart1.ChartArea.Fill.Tile = false;

        // Set the image offset.
        chart1.ChartArea.Fill.PicStretch.Left = 20;
        chart1.ChartArea.Fill.PicStretch.Top = 20;
        chart1.ChartArea.Fill.PicStretch.Right = 5;
        chart1.ChartArea.Fill.PicStretch.Bottom = 5;
 
        const outputFileName = 'SetImageOffsetOfChart-out.xlsx';
        // Save the modified workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        downloadName.value = outputFileName;
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