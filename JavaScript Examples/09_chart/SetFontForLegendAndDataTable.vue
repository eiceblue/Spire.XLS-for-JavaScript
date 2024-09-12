<template>
  <span>Click the following button to set the font for legend and datalable</span>
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
        let excelFileName = 'ChartSample1.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet from workbook
        let ws = workbook.Worksheets.get(0);
        let chart = ws.Charts.get(0);

        // Create a font with specified size and color
        let font = workbook.CreateFont();
        font.Size = 14.0;
        font.Color = wasmModule.Color.get_Red();

        // Apply the font to chart Legend
        chart.Legend.TextArea.SetFont(font);

        // Apply the font to chart DataLabel
        for (let cs of chart.Series) {
            cs.DataPoints.DefaultDataPoint.DataLabels.TextArea.SetFont(font);
        }

        // Save the modified workbook 
        const outputFile = 'SetFontForLegendAndDataTable.xlsx';
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
