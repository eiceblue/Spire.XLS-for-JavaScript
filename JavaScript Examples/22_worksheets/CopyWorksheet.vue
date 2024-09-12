<template>
  <span
    >Click the following button to copy a worksheet to another workbook</span
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
        let excelFileName1 = "ReadImages.xlsx";
        let excelFileName2 = "Sample.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName1,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );
        await wasmModule.FetchFileToVFS(
          excelFileName2,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        let sourceWorkbook = wasmModule.Workbook.Create();
        sourceWorkbook.LoadFromFile({ fileName: excelFileName1 });
        //Get the first worksheet
        let srcWorksheet = sourceWorkbook.Worksheets.get(0);

        //Create a workbook
        let targetWorkbook = wasmModule.Workbook.Create();

        //Load the target Excel document from disk
        targetWorkbook.LoadFromFile({ fileName: excelFileName2 });

        //Add a new worksheet
        let targetWorksheet = targetWorkbook.Worksheets.Add("added");

        //Copy the first worksheet of source Excel document to the new added worksheet of target Excel document
        targetWorksheet.CopyFrom(srcWorksheet);

        const outputFileName = "CopyWorksheet_output.xlsx";
        targetWorkbook.SaveToFile({ fileName: outputFileName });

        // Dispose of the workbook object to release resources
        targetWorkbook.Dispose();

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
  