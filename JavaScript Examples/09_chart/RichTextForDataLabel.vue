<template>
  <span>Click the following button to set rich text for DataPoints dataLabel in worksheet</span>
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
        let excelFileName = 'ChartToImage.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get first worksheet of the workbook
        let worksheet = workbook.Worksheets.get(0);

        //Get the first chart inside this worksheet
        let chart = worksheet.Charts.get(0);

        //Get the first datalabel of the first series
        let datalabel = chart.Series.get(0).DataPoints.get(0).DataLabels;

        //Set the text
        datalabel.Text = "Rich Text Label";

        //Show the value
        chart.Series.get(0).DataPoints.get(0).DataLabels.HasValue = true;

        //Set styles for the text
        //chart.Series.get(0).DataPoints.get(0).DataLabels.Font.Color = Module.wasmModule.Color.get_Red();
        //chart.Series.get(0).DataPoints.get(0).DataLabels.Font.IsBold = true;
        chart.Series.get(0).DataPoints.get(0).DataLabels.Color = wasmModule.Color.get_Red();
        chart.Series.get(0).DataPoints.get(0).DataLabels.IsBold = true;

        // Save the modified workbook 
        const outputFile = 'RichTextForDataLabel.xlsx';
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
