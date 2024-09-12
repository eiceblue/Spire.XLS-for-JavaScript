<template>
  <span
    >Click the following button to set font and background color for TextBox in
    Excel file</span
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
        let inputFileName = "Template_Xls_5.xlsx";
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

        //Get the textbox which will be edited.
        let shape = sheet.TextBoxes.get(0);

        //Set the font and background color for the textbox.
        //Set font.
        let font = workbook.CreateFont();
        //font.IsStrikethrough = true
        font.FontName = "Century Gothic";
        font.Size = 10;
        font.IsBold = true;
        font.Color = wasmModule.Color.get_Blue();
        let rto = shape.RichText;
        let rt = wasmModule.RichTextShape.Convert(rto);
        rt.SetFont(0, shape.Text.length - 1, font);

        //Set background color
        shape.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
        shape.Fill.ForeKnownColor = wasmModule.ExcelColors.BlueGray;

        let outputFileName = "SetFontAndBackground_output.xlsx";
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
  