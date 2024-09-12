<template>
  <span>Click the following button to autofit column width and row height based on cell value</span>
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
        let worksheet = workbook.Worksheets.get(0);

        //Set value for B8
        let cell = worksheet.Range.get("B8");
        cell.Text = "Welcome to Spire.XLS!";

        //Set the cell style
        let style = cell.Style;
        style.Font.Size = 16;
        style.Font.IsBold = true;

        //Autofit column width and row height based on cell value
        cell.AutoFitColumns();
        cell.AutoFitRows();

        // Define the output file name
        const outputFileName = "AutoFitBasedOnCellValue.xlsx";

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
