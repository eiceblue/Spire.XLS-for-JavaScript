<template>
  <span>Click the following button to detect the empty worksheet</span>
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

        let worksheet1 = workbook.Worksheets.get(0);

        //Detect the first worksheet is empty or not
        let detect1 = worksheet1.IsEmpty;

        //Get the second worksheet
        let worksheet2 = workbook.Worksheets.get(1);

        //Detect the second worksheet is empty or not
        let detect2 = worksheet2.IsEmpty;

        //Create StringBuilder to save
        let content = [];

        //Set string format for displaying
        let result = `The first worksheet is empty or not: ${detect1}\r\nThe second worksheet is empty or not: ${detect2}`;

        //Add result string to StringBuilder
        content.push(result);

        let outputFileName = "DetectEmptyWorksheet_output.txt";
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
  