<template>
  <span
    >Click the following button to apply style to an entire excel
    worksheet</span
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
      if (wasmModule.value) {
        // Load the fonts
        await wasmModule.value.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Load the file
        let excelFileName = "worksheetSample1.xlsx";
        await wasmModule.value.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        let workbook = wasmModule.value.Workbook.Create();
        workbook.LoadFromFile(excelFileName);
        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Create a cell style
        let style = workbook.Styles.Add("newStyle");
        style.Color = wasmModule.value.Color.get_LightBlue();
        style.Font.Color = wasmModule.value.Color.get_White();
        style.Font.Size = 15;
        style.Font.IsBold = true;

        //Apply the style to the first worksheet
        sheet.ApplyStyle(style);

        const outputFileName = "ApplyStyleToWorksheet_output.xlsx";
        workbook.SaveToFile({ fileName: outputFileName });

        // Dispose of the workbook object to release resources
        workbook.Dispose();

        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.value.FS.readFile(outputFileName);
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
  