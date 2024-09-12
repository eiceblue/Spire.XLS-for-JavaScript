<template>
  <span>Click the following button to add list data validation in Excel file</span>
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
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='ListDataValidation.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Set text for cells 
        sheet.Range.get("A7").Text = "Beijing";
        sheet.Range.get("A8").Text = "New York";
        sheet.Range.get("A9").Text = "Denver";
        sheet.Range.get("A10").Text = "Paris";

        //Set data validation for cell
        let range = sheet.Range.get("D10");
        range.DataValidation.ShowError = true;
        range.DataValidation.AlertStyle = wasmModule.AlertStyleType.Stop;
        range.DataValidation.ErrorTitle = "Error";
        range.DataValidation.ErrorMessage = "Please select a city from the list";
        range.DataValidation.DataRange = sheet.Range.get("A7:A10");

        const outputFileName = 'ListDataValidation-out.xlsx';
        // Save the modified workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
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
