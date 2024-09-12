<template>
  <span
    >Click the following button to detect if an Excel document contains vba
    macros</span
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
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the fonts
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Load the files
        let excelFileName = "MacroSample.xls";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile(excelFileName);

        //Detect if the Excel file contains VBA macros
        let value = [];
        let hasMacros = book.HasMacros;
        if (hasMacros) {
          value.push("Yes");
        } else {
          value.push("No");
        }
        let outputFileName = "DetectVBAMacros_output.txt";
        wasmModule.FS.writeFile(outputFileName, value.join("\n"));

        // Dispose of the workbook object to release resources
        book.Dispose();
        
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
  