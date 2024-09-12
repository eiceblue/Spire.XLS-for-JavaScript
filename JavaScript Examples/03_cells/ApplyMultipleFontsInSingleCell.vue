<template>
  <span>Click the following button to apply multiple fonts in single cell in Excel file</span>
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

        // Create a font object in workbook, setting the font color, size and type
        let font1 = workbook.CreateFont();
        font1.KnownColor = wasmModule.ExcelColors.LightBlue;
        font1.IsBold = true;
        font1.Size = 10;

        // Create another font object specifying its properties
        let font2 = workbook.CreateFont();
        font2.KnownColor = wasmModule.ExcelColors.Red;
        font2.IsBold = true;
        font2.IsItalic = true;
        font2.FontName = "Times New Roman";
        font2.Size = 11;

        // Write a RichText string to the cell 'H5', and set the font for it
        let richText = sheet.Range.get("H5").RichText;
        richText.Text = "This document was created with Spire.XLS for .NET.";
        richText.SetFont(0, 29, font1);
        richText.SetFont(31, 48, font2);

        // Define the output file name
        const outputFileName = "ApplyMultipleFontsInSingleCell.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile(outputFileName);

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
