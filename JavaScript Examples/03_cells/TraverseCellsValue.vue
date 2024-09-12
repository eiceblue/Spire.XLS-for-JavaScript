<template>
  <span
    >Click the following button to traverse cells value in worksheet</span
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

        // Load the sample file into the virtual file system (VFS)
        let excelFileName = "CellValues.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Load an existing Excel from the virtual file system
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let worksheet = workbook.Worksheets.get(0);

        // Create a string builder
        let builder = [];

        // Get the cell range collection
        let cellRangeCollection = worksheet.Cells;

        builder.push("Values of the first sheet:");

        // Traverse cells value
        for (let cellRange of cellRangeCollection) {
          // Set string format for displaying
          let result = `Cell: ${cellRange.RangeAddress}   Value: ${cellRange.Value}`;

          // Add result string to content array
          builder.push(result);
        }

        // Combine all the found data into a single string
        let content = builder.join("\n");

        // Define the output file name
        const outputFileName = "TraverseCellsValue.txt";

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
