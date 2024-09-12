<template>
  <span>Click the following button to add scrollbar control in Excel file</span>
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

        // Add an empty sheet to the workbook
        workbook.CreateEmptySheets(1);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Set a value for range B10
        sheet.Range.get("B10").NumberValue = 1;
        sheet.Range.get("B10").Style.Font.IsBold = true;

        // Add scroll bar control
        let scrollBar = sheet.ScrollBarShapes.AddScrollBar(10, 3, 150, 20);
        scrollBar.LinkedCell = sheet.Range.get("B10");
        scrollBar.Min = 1;
        scrollBar.Max = 150;
        scrollBar.IncrementalChange = 1;
        scrollBar.Display3DShading = true;

        // Define the output file name
        const outputFileName = "AddScrollBarControl.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2010,
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
