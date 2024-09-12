<template>
  <span>Click the following button to change chart series color</span>
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
        let excelFileName = 'ChangeSeriesColor.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);
        //Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        //Get the first chart
        let chart = sheet.Charts.get(0);

        //Get the second series
        let cs = chart.Series.get(1);

        //Set the fill type
        cs.Format.Fill.FillType = wasmModule.ShapeFillType.SolidColor;

        //Change the fill color
        cs.Format.Fill.ForeColor = wasmModule.Color.get_Orange();

        // Save the modified workbook 
        const outputFile = 'ChangeSeriesColor.xlsx';
        workbook.SaveToFile(outputFile);
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved Excel file from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/xlsx' });

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
