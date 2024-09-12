<template>
  <span>Click the following button to set excel worksheet page margins</span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the fonts
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Load the files
        let excelFileName = "WorksheetSample1.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = book.Worksheets.get(0);

        // Set margins for top, bottom, left and right, here the unit of measure is Inch
        sheet.PageSetup.TopMargin = 0.3;
        sheet.PageSetup.BottomMargin = 1;
        sheet.PageSetup.LeftMargin = 0.2;
        sheet.PageSetup.RightMargin = 1;
        // Set the header margin and footer margin
        sheet.PageSetup.HeaderMarginInch = 0.1;
        sheet.PageSetup.FooterMarginInch = 0.5;

        let outputFileName = "SetMargins_output.xlsx";
        //Save the document
        book.SaveToFile({ fileName: outputFileName });

        // Dispose of the workbook object to release resources
        book.Dispose();
        
        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        // download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
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
  