<template>
  <span>Click the following button to get the range of the freeze pane</span>
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
        // Load the fonts
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Load the file
        let excelFileName = "WorksheetSample2.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({ fileName: excelFileName });

        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);
        let rowIndex = null;
        let colIndex = null;
        let r = [];
        //The row and column index of the frozen pane is passed through the out parameter.
        //If it returns to 0, it means that it is not frozen
        let indexs = sheet.GetFreezePanes();
        colIndex = indexs[1];
        rowIndex = indexs[0];

        r.push(`Row index: ${rowIndex}, column index: ${colIndex}`);

        let outputFileName = "GetFreezePaneRange_output.txt";

        // Write r to txt file
        wasmModule.FS.writeFile(outputFileName, r.join("\n"));

        // Dispose of the workbook object to release resources
        workbook.Dispose();

        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "text/plain",
        });

        // download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
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
  