<template>
  <span
    >Click the following button to detect if an excel workbook is password
    protected</span
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
        let inputFileName = "ProtectedWorkbook.xlsx";
        await wasmModule.FetchFileToVFS(
          inputFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        const value = wasmModule.Workbook.IsPasswordProtected(inputFileName);
        let boolvalue = [];
        boolvalue.push(value ? "Yes" : "No");

        let outputFileName = "DetectProtection_output.txt";
        wasmModule.FS.writeFile(outputFileName, boolvalue.join("\n"));

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
  