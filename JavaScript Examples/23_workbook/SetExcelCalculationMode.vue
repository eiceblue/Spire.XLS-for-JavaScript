<template>
  <span>Click the following button to set excel calculation mode</span>
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
        let excelFileName = "CreateTable.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile(excelFileName);

        // Set excel calculation mode as Manual
        book.CalculationMode = spirexls.ExcelCalculationMode.Manual;

        let outputFileName = "SetExcelCalculationMode_output.xlsx";
        //Save the document
        book.SaveToFile({ fileName: outputFileName });

        // Dispose of the workbook object to release resources
        book.Dispose();
        
        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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
  