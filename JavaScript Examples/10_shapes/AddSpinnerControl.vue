<template>
  <span>Click the following button to add spinner control in Excel file</span>
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
        let excelFileName = 'ExcelSample_N1.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Set text for range C11
        sheet.Range.get("C11").Text = "Value:";
        sheet.Range.get("C11").Style.Font.IsBold = true;

        // Set value for range C12
        sheet.Range.get("C12").Value2 = wasmModule.Int32.Create(0);

        // Add spinner control
        let spinner = sheet.SpinnerShapes.AddSpinner(12, 4, 20, 20);
        spinner.LinkedCell = sheet.Range.get("C12");
        spinner.Min = 0;
        spinner.Max = 100;
        spinner.IncrementalChange = 5;
        spinner.Display3DShading = true;

        // Save the modified workbook  
        const outputFile = 'AddSpinnerControl.xlsx';
        workbook.SaveToFile(outputFile);
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved Excel file from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Download the converted Excel file
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
