<template>
  <span>Click the following button to wrap or unwrap the text in Excel file</span>
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

        // Wrap the excel text
        sheet.Range.get("C1").Text =
          "e-iceblue is in facebook and welcome to like us";
        sheet.Range.get("C1").Style.WrapText = true;
        sheet.Range.get("D1").Text =
          "e-iceblue is in twitter and welcome to follow us";
        sheet.Range.get("D1").Style.WrapText = true;

        // Unwrap the excel text
        sheet.Range.get("C2").Text =
          "http://www.facebook.com/pages/e-iceblue/139657096082266";
        sheet.Range.get("C2").Style.WrapText = false;
        sheet.Range.get("D2").Text = "https://twitter.com/eiceblue";
        sheet.Range.get("D2").Style.WrapText = false;

        // Set the text color of Range["C1:D1"]
        sheet.Range.get("C1:D1").Style.Font.Size = 15;
        sheet.Range.get("C1:D1").Style.Font.Color = wasmModule.Color.get_Blue();

        // Set the text color of Range["C2:D2"]
        sheet.Range.get("C2:D2").Style.Font.Size = 15;
        sheet.Range.get("C2:D2").Style.Font.Color =
          wasmModule.Color.get_DeepSkyBlue();

        // Define the output file name
        const outputFileName = "WrapOrUnwrapTextInCells.xlsx";

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
