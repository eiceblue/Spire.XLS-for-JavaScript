<template>
  <span>Click the following button to set colors and palette for workbook</span>
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
      
        //Adding Orchid color to the paconstte at 60th index
        workbook.ChangePaletteColor(wasmModule.Color.get_Orchid(), 60);
      
        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);
      
        const cell = sheet.Range.get("B2");
        cell.Text = "Welcome to use Spire.XLS";
      
        //Set the Orchid (custom) color to the font
        cell.Style.Font.Color = wasmModule.Color.get_Orchid();
        cell.Style.Font.Size = 20;
        cell.AutoFitColumns();
        cell.AutoFitRows();

        //Save result file
        const outputFileName = 'ColorsAndPalette_out.xlsx';
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
