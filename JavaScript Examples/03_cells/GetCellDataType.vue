<template>
  <span>Click the following button to get the cell data type</span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Load the sample file into the virtual file system (VFS)
        let excelFileName = "Template_Xls_2.xlsx";
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

        // Get the cell types of the cells in range "C13:F13"
        for (let range of sheet.Range.get("H2:H7").Cells) {
          let cellType = sheet.GetCellType(range.Row, range.Column, false);
          sheet.get({ row: range.Row, column: range.Column + 1 }).Text =
            cellType.toString();
          sheet.get({
            row: range.Row,
            column: range.Column + 1,
          }).Style.Font.Color = wasmModule.Color.get_Red();
          sheet.get({
            row: range.Row,
            column: range.Column + 1,
          }).Style.Font.IsBold = true;
        }

        // Define the output file name
        const outputFileName = "GetCellDataType.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile(outputFileName);

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
