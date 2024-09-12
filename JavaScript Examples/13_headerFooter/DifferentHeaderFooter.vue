<template>
  <span>Click the following button to set different header and footer</span>
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
        let excelFileName='DifferentHeaderFooter.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);

        // Set text in range
        sheet.Range.get("A1").Text = "Page 1";
        sheet.Range.get("G1").Text = "Page 2";

        // Set the different header footer for Odd and Even pages
        sheet.PageSetup.DifferentOddEven = 1;

        // Set the header with font, size, bold, and color
        sheet.PageSetup.OddHeaderString = "&\"Arial\"&12&B&KFFC000 Odd_Header";
        sheet.PageSetup.OddFooterString = "&\"Arial\"&12&B&KFFC000 Odd_Footer";
        sheet.PageSetup.EvenHeaderString = "&\"Arial\"&12&B&KFF0000 Even_Header";
        sheet.PageSetup.EvenFooterString = "&\"Arial\"&12&B&KFF0000 Even_Footer";

        sheet.ViewMode = wasmModule.ViewMode.Layout;

        //Save result file
        const outputFileName = 'DifferentHeaderFooter_out.xlsx';
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
