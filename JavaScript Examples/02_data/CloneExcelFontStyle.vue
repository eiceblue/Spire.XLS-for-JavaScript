<template>
  <span>Click the following button to clone the font style in Excel file</span>
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
      const wasmModule = window.wasmModule; 
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

        // Add the text to the Excel sheet cell range A1
        sheet.Range.get("A1").Text = "Text1";

        // Set A1 cell range's CellStyle
        let style = workbook.Styles.Add("style");
        style.Font.FontName = "Calibri";
        style.Font.Color = wasmModule.Color.get_Red();
        style.Font.Size = 12;
        style.Font.IsBold = true;
        style.Font.IsItalic = true;
        sheet.Range.get("A1").CellStyleName = style.Name;

        // Clone the same style for B2 cell range
        let csOrieign = style.clone();
        sheet.Range.get("B2").Text = "Text2";
        sheet.Range.get("B2").CellStyleName = csOrieign.Name;

        // Clone the same style for C3 cell range and then reset the font color for the text
        let csGreen = style.clone();
        csGreen.Font.Color = wasmModule.Color.get_Green();
        sheet.Range.get("C3").Text = "Text3";
        sheet.Range.get("C3").CellStyleName = csGreen.Name;

        // Define the output file name
        const outputFileName = "CloneExcelFontStyle.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2013,
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
