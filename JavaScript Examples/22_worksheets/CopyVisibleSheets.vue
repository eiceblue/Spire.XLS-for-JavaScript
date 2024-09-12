<template>
  <span
    >Click the following button to copy only visible sheets from one workbook to
    another</span
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
        let excelFileName = "CopyVisibleSheets.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({ fileName: excelFileName });

        //Create a new workbook and clear the sheets
        let workbookNew = wasmModule.Workbook.Create();
        workbookNew.Version = spirexls.ExcelVersion.Version2013;
        workbookNew.Worksheets.Clear();

        //Loop through the worksheets
        for (let sheet of workbook.Worksheets) {
          //Judge if the worksheet is visible
          if (sheet.Visibility === spirexls.WorksheetVisibility.Visible) {
            //Copy the sheet to new workbook
            let name = sheet.Name;
            workbookNew.Worksheets.AddCopy({ sheet: sheet });
          }
        }

        const outputFileName = "CopyVisibleSheets_output.xlsx";
        workbook.SaveToFile({ fileName: outputFileName });

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
  