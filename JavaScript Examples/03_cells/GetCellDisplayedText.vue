<template>
  <span
    >Click the following button to get cell value and displayed text in worksheet</span
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
        let worksheet = workbook.Worksheets.get(0);

        // Create a string builder
        let builder = [];

        // Set value for B8
        let cell = worksheet.Range.get("B8");
        cell.NumberValue = 0.012345;

        // Set the cell style
        let style = cell.Style;
        style.NumberFormat = "0.00";

        // Get the cell value
        let cellValue = cell.Value;

        // Get the displayed text of the cell
        let displayedText = cell.DisplayedText;

        // Set string format for displaying
        let result = `B8 Value: ${cellValue}\r\nB8 displayed text: ${displayedText}`;

        // Add result string to StringBuilder
        builder.push(result);

        // Combine all the found data into a single string
        let content = builder.join("\n");

        // Define the output file name
        const outputFileName = "GetCellDisplayedText.txt";

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
