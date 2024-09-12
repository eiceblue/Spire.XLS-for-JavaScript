<template>
  <span>Click the following button to set the workbook theme</span>
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
        let inputFileName = "SetTheme.xlsx";
        await wasmModule.FetchFileToVFS(
          inputFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a workbook
        let srcWorkbook = wasmModule.Workbook.Create();
        // Load an excel file
        srcWorkbook.LoadFromFile(inputFileName);
        let srcWorksheet = srcWorkbook.Worksheets.get(0);

        let workbook = wasmModule.Workbook.Create();
        workbook.Worksheets.Clear();
        workbook.Worksheets.AddCopy({ sheet: srcWorksheet });

        // 1. Copy the theme of the workbook
        // workbook.CopyTheme(srcWorkbook);

        // 2. Set a certain type of color of the default theme in the workbook
        workbook.SetThemeColor(
          wasmModule.ThemeColorType.Dk1,
          wasmModule.Color.get_SkyBlue()
        );

        let outputFileName = "SetTheme_output.xlsx";
        // Save the document
        workbook.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.ExcelVersion.Version2013,
        });
        // Dispose of the workbook object to release resources
        workbook.Dispose();
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
  