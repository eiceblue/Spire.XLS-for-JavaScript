<template>
  <span>Click the following button to create formula and apply conditional formatting in Excel file</span>
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
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Input file
        let excelFileName='ConditionallyFormatDate.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first worksheet and the first column from the workbook.
        const sheet = workbook.Worksheets.get(0);
        const range = sheet.Columns.get(0);
      
        //Set the conditional formatting formula and apply the rule to the chosen cell range.
        const xcfs = sheet.ConditionalFormats.Add();
        xcfs.AddRange(range);
        const conditional = xcfs.AddCondition();
        conditional.FormatType = wasmModule.ConditionalFormatType.Formula;
        conditional.FirstFormula = "=($A1<$B1)";
        conditional.BackKnownColor = wasmModule.ExcelColors.Yellow;

        //Save result file
        const outputFileName = 'CreateFormulaConditionalFormat_out.xlsx';
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
