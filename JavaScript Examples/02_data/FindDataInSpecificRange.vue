<template>
  <span
    >Click the following button to find data in specific range</span
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
        let excelFileName = "FindCellsSample.xlsx";
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
        let sheet = workbook.Worksheets.get(0);

        // Specify a range
        let range = sheet.Range.get({
          row: 1,
          column: 1,
          lastRow: 12,
          lastColumn: 8,
        });

        // Create a string builder
        let builder = [];

        // Find text from this range
        FindTextFromRange(range, builder);

        // Find number from this range
        FindNumberFromRange(range, builder);

        // Combine all the found data into a single string
        let content = builder.join("\n");

        // Define the output file name
        const outputFileName = "FindDataInSpecificRange.txt";

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
    // Finds text values in the given range and appends them to the builder array
    const FindTextFromRange = (range, builder) => {
      // Find string from this range
      let textRanges = range.FindAllString("E-iceblue", false, false);

      // Append the address of found cells in builder
      if (textRanges.Count !== 0) {
        for (let r of textRanges) {
          let address = r.RangeAddress;
          builder.push("The address of found text cell is: " + address);
        }
      } else {
        builder.push("No cell contain the text");
      }
    };
    // Finds number in the given range and appends them to the builder array
    const FindNumberFromRange = (range, builder) => {
      // Find number from this range
      let numberRanges = range.FindAllNumber(100, true);

      // Append the address of found cells in builder
      if (numberRanges.Count !== 0) {
        for (let r of numberRanges) {
          let address = r.RangeAddress;
          builder.push("The address of found number cell is: " + address);
        }
      } else {
        builder.push("No cell contain the number");
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
