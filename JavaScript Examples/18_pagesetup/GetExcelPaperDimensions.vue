<template>
  <span
    >The example demonstrates how to get the dimensions of Excel paper.</span
  >
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"
    >Click here to download the generated file</a
  >
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
        // Load the font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Create a new workbook
        const book = wasmModule.Workbook.Create();
        // Get the first worksheet.
        let sheet = book.Worksheets.get(0);
        let content = [];
        // Get the dimensions of A2 paper.
        sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.A2Paper;
        content.push(
          "A2Paper: " +
            sheet.PageSetup.PageWidth +
            " x " +
            sheet.PageSetup.PageHeight
        );

        // Get the dimensions of A3 paper.
        sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperA3;
        content.push(
          "PaperA3: " +
            sheet.PageSetup.PageWidth +
            " x " +
            sheet.PageSetup.PageHeight
        );

        // Get the dimensions of A4 paper.
        sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperA4;
        content.push(
          "PaperA4: " +
            sheet.PageSetup.PageWidth +
            " x " +
            sheet.PageSetup.PageHeight
        );

        // Get the dimensions of paper letter.
        sheet.PageSetup.PaperSize = wasmModule.PaperSizeType.PaperLetter;
        content.push(
          "PaperLetter: " +
            sheet.PageSetup.PageWidth +
            " x " +
            sheet.PageSetup.PageHeight
        );
        // Define the output file name
        const outputFileName = "GetExcelPaperDimensions.txt";
        // Save result file
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "text/plain",
        });

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);

        // Clean up resources
        book.Dispose();
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
