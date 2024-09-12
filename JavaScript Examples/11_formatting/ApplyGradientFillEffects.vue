<template>
  <span>Click the following button to apply gradient filling effects</span>
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

        //Create a workbook
        const workbook = wasmModule.Workbook.Create();
        workbook.Version = wasmModule.ExcelVersion.Version2010;

        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);
        //Get "B5" cell
        const range = sheet.Range.get("B5");

        //Set row height and column width
        range.RowHeight = 50;
        range.ColumnWidth = 30;
        range.Text = "Hello";

        //Set alignment style
        range.Style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;

        //Set gradient filling effects
        range.Style.Interior.FillPattern = wasmModule.ExcelPatternType.Gradient;
        range.Style.Interior.Gradient.ForeColor = wasmModule.Color.FromArgb(255, 255, 255);
        range.Style.Interior.Gradient.BackColor = wasmModule.Color.FromArgb(79, 129, 189);
        range.Style.Interior.Gradient.TwoColorGradient(wasmModule.GradientVariantsType.HorizontalAlignTypeHorizontal, wasmModule.GradientVariantsType.Color);

        //Save result file
        const outputFileName = 'ApplyGradientFillEffects_out.xlsx';
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
