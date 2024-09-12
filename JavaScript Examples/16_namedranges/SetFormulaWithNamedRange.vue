<template>
  <span
    >The example demonstrates how to set formula with named range in Excel
    file</span
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

        // Input file
        let excelFileName = "ExcelSample_N1.xlsx";
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
        // Get the first worksheet
        let sheet = book.Worksheets.get(0);

        // Create a named range
        let namedRange = book.NameRanges.Add("MyNamedRange");

        // Refers to range
        namedRange.RefersToRange = sheet.Range.get("B10:B12");

        // Set the formula of range to named range
        sheet.Range.get("B13").Formula = "=SUM(MyNamedRange)";

        // Set value of ranges
        sheet.Range.get("B10").Value2 = wasmModule.Int32.Create(10);
        sheet.Range.get("B11").Value2 = wasmModule.Int32.Create(20);
        sheet.Range.get("B12").Value2 = wasmModule.Int32.Create(30);

        // Define the output file name
        const outputFileName = "SetFormulaWithNamedRange.xlsx";
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
