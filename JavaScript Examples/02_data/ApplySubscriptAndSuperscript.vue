<template>
  <span>Click the following button to apply subscript and superscript in Excel file</span>
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
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        //Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        sheet.Range.get("B2").Text = "This is an example of Subscript:";
        sheet.Range.get("D2").Text = "This is an example of Superscript:";

        // Set the rtf value of "B3" to "R100-0.06"
        let range = sheet.Range.get("B3");
        range.RichText.Text = "R100-0.06";

        // Create a font. Set the IsSubscript property of the font to "true"
        let font = workbook.CreateFont();
        font.IsSubscript = true;
        font.Color = wasmModule.Color.get_Green();

        // Set font for specified range of the text in "B3"
        range.RichText.SetFont(4, 8, font);

        // Set the rtf value of "D3" to "a2 + b2 = c2"
        range = sheet.Range.get("D3");
        range.RichText.Text = "a2 + b2 = c2";

        // Create a font. Set the IsSuperscript property of the font to "true"
        font = workbook.CreateFont();
        font.IsSuperscript = true;

        // Set font for specified range of the text in "D3"
        range.RichText.SetFont(1, 1, font);
        range.RichText.SetFont(6, 6, font);
        range.RichText.SetFont(11, 11, font);

        sheet.AllocatedRange.AutoFitColumns();

        // Define the output file name
        const outputFileName = "ApplySubscriptAndSuperscript.xlsx";

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
