<template>
  <span
    >Click the following button to set internal margin of the text box in Excel
    file</span
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
        let inputFileName = "Template_Xls_4.xlsx";
        await wasmModule.FetchFileToVFS(
          inputFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile(inputFileName);

        //Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        //Add a textbox to the sheet and set its position and size.
        let textbox = sheet.TextBoxes.AddTextBox(4, 2, 100, 300);

        //Set the text on the textbox.
        textbox.Text =
          "Insert TextBox in Excel and set the margin for the text";
        textbox.HAlignment = wasmModule.CommentHAlignType.Center;
        textbox.VAlignment = wasmModule.CommentVAlignType.Center;

        //Set the inner margins of the contents.
        textbox.InnerLeftMargin = 1;
        textbox.InnerRightMargin = 3;
        textbox.InnerTopMargin = 1;
        textbox.InnerBottomMargin = 1;

        let outputFileName = "SetInternalMarginOfTextbox_output.xlsx";
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
  },
};
</script>
  