<template>
  <span
    >Click the following button to write RichText into a cell in Excel
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
      const wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Create an underlined font
        let fontUnderline = workbook.CreateFont();
        fontUnderline.Underline = wasmModule.FontUnderlineType.Single;

        // Create an italic font
        let fontItalic = workbook.CreateFont();
        fontItalic.IsItalic = true;

        // Create a green-colored font
        let fontColor = workbook.CreateFont();
        fontColor.KnownColor = wasmModule.ExcelColors.Green;

        // Get the rich text object for cell B11
        let richText = sheet.Range.get("B11").RichText;
        richText.Text = "Bold and underlined and italic and colored text.";

        // Apply the bold font to the range from character 0 to 3 (inclusive)
        richText.SetFont(0, 3, fontBold);

        // Apply the underline font to the range from character 9 to 18 (inclusive)
        richText.SetFont(9, 18, fontUnderline);

        // Apply the italic font to the range from character 24 to 29 (inclusive)
        richText.SetFont(24, 29, fontItalic);

        // Apply the green color font to the range from character 35 to 41 (inclusive)
        richText.SetFont(35, 41, fontColor);

        // Define the output file name
        const outputFileName = "WriteRichText_out.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);

        // Clean up resources
        workbook.Dispose();
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
