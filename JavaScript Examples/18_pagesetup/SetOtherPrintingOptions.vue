<template>
  <span>
    The example demonstrates how to set other printing options of Excel file.
  </span>
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
        // Load the arial.ttf font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );
        // Input file
        let excelFileName = "Template_Xls_4.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile({
          fileName: excelFileName,
          version: wasmModule.ExcelVersion.Version2010,
        });
        // Get the first worksheet.
        let sheet = book.Worksheets.get(0);

        // Get the reference of the PageSetup of the worksheet.
        let pageSetup = sheet.PageSetup;

        // Allow to print gridlines.
        pageSetup.IsPrintGridlines = true;

        // Allow to print row/column headings.
        pageSetup.IsPrintHeadings = true;

        // Allow to print worksheet in black & white mode.
        pageSetup.BlackAndWhite = true;

        // Allow to print comments as displayed on worksheet.
        pageSetup.PrintComments = wasmModule.PrintCommentType.InPlace;

        // Allow to print worksheet with draft quality.
        pageSetup.Draft = true;

        // Allow to print cell errors as N/A.
        pageSetup.PrintErrors = wasmModule.PrintErrorsType.NA;
        // Define the output file name
        const outputFileName = "SetOtherPrintingOptions.xlsx";
        // Save the workbook to the specified path
        book.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2010,
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