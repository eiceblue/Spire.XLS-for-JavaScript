<template>
  <span>Click the following button to set the border color and styles for chart </span>
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
        let excelFileName = 'ChartSample3.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 

        workbook.LoadFromFile(excelFileName);

        //Get the first worksheet from workbook and then get the first chart from the worksheet
        let ws = workbook.Worksheets.get(0);
        let chart = ws.Charts.get(0);

        //Set CustomLineWeight property for Series line
        chart.Series.get(0).DataPoints.get(0).DataFormat.LineProperties.CustomLineWeight = 2.5;
        //Set color property for Series line
        chart.Series.get(0).DataPoints.get(0).DataFormat.LineProperties.Color = wasmModule.Color.get_Red();

        // Save the modified workbook 
        const outputFile = 'SetBorderColorAndStyle.xlsx';
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
