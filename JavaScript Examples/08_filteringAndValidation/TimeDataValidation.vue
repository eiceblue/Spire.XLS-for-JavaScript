<template>
  <span>Click the following button to add time data validation in Excel file</span>
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

        let inputFileName='TimeDataValidation.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        sheet.Range.get("C12").Text = "Please enter time between 09:00 and 18:00:";
        sheet.Range.get("C12").AutoFitColumns();

        //Set Time data validation for cell "D12"
        let range = sheet.Range.get("D12");
        range.DataValidation.AllowType = wasmModule.CellDataType.Time;
        range.DataValidation.CompareOperator = wasmModule.ValidationComparisonOperator.Between;

        range.DataValidation.Formula1 = "09:00";
        range.DataValidation.Formula2 = "18:00";

        range.DataValidation.AlertStyle = wasmModule.AlertStyleType.Info;
        range.DataValidation.ShowError = true;
        range.DataValidation.ErrorTitle = "Time Error";
        range.DataValidation.ErrorMessage = "Please enter a valid time";
        range.DataValidation.InputMessage = "Time Validation Type";
        range.DataValidation.IgnoreBlank = true;
        range.DataValidation.ShowInput = true;
        
        const outputFileName = 'TimeDataValidation-out.xlsx';
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
