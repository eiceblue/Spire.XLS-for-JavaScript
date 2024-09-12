<template>
  <span
    >Click the following button to find string and number in Excel file</span
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

        //Find cells with the input string
        let textRanges = sheet.FindAllString("E-iceblue", false, false);
        
        // Create a string builder
        let builder = [];

        // Append the address of found cells in builder
        if (textRanges.length !== 0) {
          for (let range of textRanges) {
            let address = range.RangeAddress;
            builder.push("The address of found text cell is: " + address);
          }
        } else {
          builder.push("No cells that contain the text");
        }

        // Find cells with the input integer or double
        let numberRanges = sheet.FindAllNumber(100, true);

        // Append the address of found cells in builder
        if (numberRanges.length !== 0) {
          for (let range of numberRanges) {
            let address = range.RangeAddress;
            builder.push("The address of found number cell is: " + address);
          }
        } else {
          builder.push("No cells that contain the number");
        }

        // Combine all the found data into a single string
        let content = builder.join("\n");

        // Define the output file name
        const outputFileName = "FindStringAndNumber.txt";

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
