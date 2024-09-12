<template>
  <span
    >Click the following button to set DB Num formatting in Excel
    file</span
  >
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

        workbook.CreateEmptySheets(1);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Set value for cells
        sheet.Range.get("A1").Value2 = wasmModule.Int32.Create(123);
        sheet.Range.get("A2").Value2 = wasmModule.Int32.Create(456);
        sheet.Range.get("A3").Value2 = wasmModule.Int32.Create(789);

        // Get the cell range
        let range = sheet.Range.get("A1:A3");

        // Set the DB num format
        range.NumberFormat = "[DBNum2][$-804]General";

        // Auto fit columns
        range.AutoFitColumns();

        // Define the output file name
        const outputFileName = "SetDBNumFormatting.xlsx";

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
