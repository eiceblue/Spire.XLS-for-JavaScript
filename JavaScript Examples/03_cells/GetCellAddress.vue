<template>
  <span
    >Click the following button to get cell address in Excel file</span
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

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Create a string builder
        let builder = [];

        // Get a cell range
        let range = sheet.Range.get("A1:B5");

        // Get address of range
        let address = range.RangeAddressLocal;
        builder.push(`Address of range: ${address}`);

        // Get the cell count of range
        let count = range.CellsCount;
        builder.push(`Cell count of range: ${count}`);

        // Get the address of the entire column of range
        let entireColAddress = range.EntireColumn.RangeAddressLocal;
        builder.push(
          `Address of entire column of the range: ${entireColAddress}`
        );

        // Get the address of the entire row of range
        let entireRowAddress = range.EntireRow.RangeAddressLocal;
        builder.push(`Address of entire row of the range ${entireRowAddress}`);

        // Combine all the found data into a single string
        let content = builder.join("\n");

        // Define the output file name
        const outputFileName = "GetCellAddress.txt";

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, content);

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "text/plain",
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
