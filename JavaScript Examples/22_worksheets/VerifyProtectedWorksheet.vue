<template>
  <span
    >Click the following button to verify if a worksheet is password
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

        // Load the file
        let excelFileName = "ProtectedWorksheet.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({ fileName: excelFileName });
        //Get the first worksheet
        let worksheet = workbook.Worksheets.get(0);

        //Verify the first worksheet
        let detect = worksheet.IsPasswordProtected;

        //Create StringBuilder to save
        let content = [];

        //Set string format for displaying
        let result = `The first worksheet is password protected or not: ${detect}`;

        //Add result string to StringBuilder
        content.push(result);

        const outputFileName = "VerifyProtectedWorksheet_ouptut.txt";
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));

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
  