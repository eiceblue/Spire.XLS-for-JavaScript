<template>
  <span
    >Click the following button to retrieve data from one Excel and extract to a new Excel file</span
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
        let excelFileName = "Template_Xls_3.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook instance and get the first worksheet.
        let newBook = wasmModule.Workbook.Create();
        let newSheet = newBook.Worksheets.get(0);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Load an existing Excel from the virtual file system
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Retrieve data and extract it to the first worksheet of the new excel workbook
        let i = 1;
        let columnCount = sheet.Columns.Count;
        let cells = sheet.Columns.get(0).Cells;
        for (let j = 0; j < cells.Count; j++) {
          let range = cells.get(j);
          if (range.Text === "teacher") {
            let sourceRange = sheet.Range.get({
              row: range.Row,
              column: 1,
              lastRow: range.Row,
              lastColumn: columnCount,
            });
            let destRange = newSheet.Range.get({
              row: i,
              column: 1,
              lastRow: i,
              lastColumn: columnCount,
            });
            sheet.Copy({
              sourceRange: sourceRange,
              destRange: destRange,
              copyStyle: true,
            });
            i += 1;
          }
        }

        // Define the output file name
        const outputFileName = "RetrieveAndExtractData.xlsx";

        // Save the workbook to the specified path
        newBook.SaveToFile({
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
        newBook.Dispose();
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
