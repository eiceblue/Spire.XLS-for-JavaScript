<template>
  <span>Click the following button to set ConditionalFormat formula in worksheet</span>
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
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        //Create a workbook
        const workbook = wasmModule.Workbook.Create();

        //Get the default first worksheet
        const sheet = workbook.Worksheets.get(0);

        //Add ConditionalFormat
        const xcfs = sheet.ConditionalFormats.Add();

        //Define the range
        xcfs.AddRange(sheet.Range.get("B5"));

        //Add condition
        const format = xcfs.AddCondition();
        format.FormatType = wasmModule.ConditionalFormatType.CellValue;

        //If greater than 1000
        format.FirstFormula = "1000";
        format.Operator = wasmModule.ComparisonOperatorType.Greater;
        format.BackColor = wasmModule.Color.get_Orange();

        sheet.get("B1").NumberValue = 40;
        sheet.get("B2").NumberValue = 500;
        sheet.get("B3").NumberValue = 300;
        sheet.get("B4").NumberValue = 400;

        //Set a SUM formula for B5
        sheet.get("B5").Formula = "=SUM(B1:B4)";

        //Add text
        sheet.get("C5").Text = "If Sum of B1:B4 is greater than 1000, B5 will have orange background.";

        //Save result file
        const outputFileName = 'SetConditionalFormatFormula_out.xlsx';
        workbook.SaveToFile({fileName: outputFileName, version:wasmModule.ExcelVersion.Version2010});

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        // Download the result file
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
