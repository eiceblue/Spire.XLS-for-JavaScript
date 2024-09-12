<template>
  <span
    >Click the following button to add data validation in Excel file</span
  >
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      const wasmModule = window.wasmModule; 
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        
        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Decimal DataValidation
        sheet.Range.get("B11").Text = "Input Number(3-6):";
        let rangeNumber = sheet.Range.get("B12");
        rangeNumber.DataValidation.CompareOperator =
          wasmModule.ValidationComparisonOperator.Between;
        rangeNumber.DataValidation.Formula1 = "3";
        rangeNumber.DataValidation.Formula2 = "6";
        rangeNumber.DataValidation.AllowType =
          wasmModule.CellDataType.Decimal;
        rangeNumber.DataValidation.ErrorMessage =
          "Please input correct number!";
        rangeNumber.DataValidation.ShowError = true;
        rangeNumber.Style.KnownColor =
          wasmModule.ExcelColors.Gray25Percent;

        // Date DataValidation
        sheet.Range.get("B14").Text = "Input Date:";
        let rangeDate = sheet.Range.get("B15");
        rangeDate.DataValidation.AllowType = wasmModule.CellDataType.Date;
        rangeDate.DataValidation.CompareOperator =
          wasmModule.ValidationComparisonOperator.Between;
        rangeDate.DataValidation.Formula1 = "1/1/1970";
        rangeDate.DataValidation.Formula2 = "12/31/1970";
        rangeDate.DataValidation.ErrorMessage = "Please input correct date!";
        rangeDate.DataValidation.ShowError = true;
        rangeDate.DataValidation.AlertStyle =
          wasmModule.AlertStyleType.Warning;
        rangeDate.Style.KnownColor = wasmModule.ExcelColors.Gray25Percent;

        // TextLength DataValidation
        sheet.Range.get("B17").Text = "Input Text:";
        let rangeTextLength = sheet.Range.get("B18");
        rangeTextLength.DataValidation.AllowType =
          wasmModule.CellDataType.TextLength;
        rangeTextLength.DataValidation.CompareOperator =
          wasmModule.ValidationComparisonOperator.LessOrEqual;
        rangeTextLength.DataValidation.Formula1 = "5";
        rangeTextLength.DataValidation.ErrorMessage = "Enter a Valid String!";
        rangeTextLength.DataValidation.ShowError = true;
        rangeTextLength.DataValidation.AlertStyle =
          wasmModule.AlertStyleType.Stop;
        rangeTextLength.Style.KnownColor =
          wasmModule.ExcelColors.Gray25Percent;

        sheet.AutoFitColumn(2);

        // Define the output file name
        const outputFileName = "DataValidation_out.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);

        // Clean up resources
        workbook.Dispose();
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
