<template>
  <span>Click the following button to get list of the fonts used</span>
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
        let excelFileName = "templateAz.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({ fileName: excelFileName });

        let fonts = [];

        // Loop all sheets of workbook
        for (let i = 0; i < workbook.Worksheets.Count; i++) {
          let sheet = workbook.Worksheets.get(i);
          for (let r = 0; r < sheet.Rows.Count; r++) {
            for (let c = 0; c < sheet.Rows.get(r).Cells.Count; c++) {
              // Get the font of cell and add it to list
              let cell = sheet.Rows.get(r).Cells.get(c);
              fonts.push(cell.Style.Font);
            }
          }
        }
        let strB = [];

        for (let font of fonts) {
          strB.push(`FontName:${font.FontName}; FontSize:${font.Size}`);
        }

        let outputFileName = "GetListOfFontsUsed_output.txt";
        wasmModule.FS.writeFile(outputFileName, strB.join("\n"));

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
  