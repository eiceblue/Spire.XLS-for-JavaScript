<template>
  <span>Click the following button to use predefined styles</span>
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

        // Create a workbook
        const workbook = wasmModule.Workbook.Create();

        // Get the first sheet
        const sheet = workbook.Worksheets.get(0);

        // Create a new style
        const style = workbook.Styles.Add("newStyle");
        style.Font.FontName = "Calibri";
        style.Font.IsBold = true;
        style.Font.Size = 15;
        style.Font.Color = wasmModule.Color.get_CornflowerBlue();

        // Get "B5" cell
        const range = sheet.Range.get("B5");
        range.Text = "Welcome to use Spire.XLS";
        range.CellStyleName = style.Name;
        range.AutoFitColumns();

        //Save result file
        const outputFileName = 'UsePredefinedStyles_out.xlsx';
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
