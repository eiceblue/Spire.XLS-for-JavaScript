<template>
  <span
    >Click the following button to replace text in textbox in Excel file</span
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

        // Load the files
        let inputFileName = "FormulasSample.xlsx";
        await wasmModule.FetchFileToVFS(
          inputFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile(inputFileName);

        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        const tag = "TAG_1$TAG_2";
        const replace = "Spire.XLS for .NET$Spire.XLS for JAVA";

        let tags = tag.split("$");
        let replacements = replace.split("$");

        for (let i = 0; i < tags.length; i++) {
          //Replace text in textbox
          _ReplaceTextInTextBox(sheet, `<${tags[i]}>`, replacements[i]);
        }

        let outputFileName = "ReplaceTextInTextBox_output.xlsx";
        //Save the document
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

    function _ReplaceTextInTextBox(sheet, sFind, sReplace) {
      // Get the textboxes of sheet
      let textBoxes = sheet.TextBoxes;
      // replace text in each textbox
      for (let tb of textBoxes) {
        if (tb.Text && tb.Text.includes(sFind)) {
          tb.Text = tb.Text.replace(sFind, sReplace);
        }
      }
    }
  },
};
</script>
  