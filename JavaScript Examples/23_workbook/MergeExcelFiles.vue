<template>
  <span>Click the following button to merge excel files</span>
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
        const files = [
          "MergeExcelFiles-1.xlsx",
          "MergeExcelFiles-2.xls",
          "MergeExcelFiles-3.xlsx",
        ];
        for (const file of files) {
          await wasmModule.FetchFileToVFS(
            file,
            "",
            `${import.meta.env.BASE_URL}static/data/`
          );
        }

        let newbook = wasmModule.Workbook.Create();
        newbook.Version = wasmModule.ExcelVersion.Version2013;
        // Clear all worksheets
        newbook.Worksheets.Clear();

        let tempbook = wasmModule.Workbook.Create();

        for (const file of files) {
          // Load the file
          tempbook.LoadFromFile(file.split("/").pop());

          for (let i = 0; i < tempbook.Worksheets.Count; i++) {
            let sheet = tempbook.Worksheets.get(i);
            // Copy every sheet in a workbook
            wasmModule.XlsWorksheetsCollection.Convert(
              newbook.Worksheets
            ).AddCopy({
              sheet: sheet,
              flags: wasmModule.WorksheetCopyType.CopyAll,
            });
          }
        }

        let outputFileName = "MergeExcelFiles_output.xlsx";
        //Save the document
        newbook.SaveToFile({ fileName: outputFileName });

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
  