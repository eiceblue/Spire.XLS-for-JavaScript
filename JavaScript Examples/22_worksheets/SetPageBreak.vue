<template>
  <span
    >Click the following button to set horizontal and vertical page break in
    excel workbook</span
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
        let excelFileName = "WorksheetSample1.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({ fileName: excelFileName });

        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Set Excel Page Break Horizontally
        sheet.HPageBreaks.Add(sheet.Range.get("A8"));
        sheet.HPageBreaks.Add(sheet.Range.get("A14"));

        //Set Excel Page Break Vertically
        //sheet.VPageBreaks.Add(sheet.Range.get("B1"));
        //sheet.VPageBreaks.Add(sheet.Range.get("C1"));

        //Set view mode to Preview mode
        workbook.Worksheets.get(0).ViewMode = spirexls.ViewMode.Preview;

        const outputFileName = "SetPageBreak_output.xlsx";
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
  