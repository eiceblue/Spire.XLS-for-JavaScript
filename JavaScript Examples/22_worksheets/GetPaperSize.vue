<template>
  <span>Click the following button to get paper size of worksheet</span>
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

        // loop the worksheet and get the PageSetup of sheet
        let sb = [];
        for (let i = 0; i < workbook.Worksheets.Count; i++) {
          let sheet = workbook.Worksheets.get(i);
          let width = sheet.PageSetup.PageWidth;
          let height = sheet.PageSetup.PageHeight;
          sb.push(sheet.Name);
          sb.push(`Width: ${width}\tHeight: ${height}\r\n`);
        }

        let outputFileName = "GetWorksheetNames_output.txt";
        // Write sb to txt file
        wasmModule.FS.writeFile(outputFileName, sb.join("\n"));

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
  