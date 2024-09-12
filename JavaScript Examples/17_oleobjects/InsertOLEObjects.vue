<template>
  <span>The example demonstrates how to insert OLE objects</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"
    >Click here to download the generated file</a
  >
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
        // Load the font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );
        // Input file
        let excelFileName = "InsertOLEObjects.xls";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Get the first worksheet
        let ws = workbook.Worksheets.get(0);
        ws.Range.get("A1").Text = "Here is an OLE Object.";

        // Insert OLE object
        let book = wasmModule.Workbook.Create();
        book.LoadFromFile(excelFileName);
        let worksheet = book.Worksheets.get(0);
        worksheet.PageSetup.LeftMargin = 0;
        worksheet.PageSetup.RightMargin = 0;
        worksheet.PageSetup.TopMargin = 0;
        worksheet.PageSetup.BottomMargin = 0;
        // Convert worksheet to image
        let image = worksheet.ToImage(1, 1, 19, 5);
        // Clean up resources
        book.Dispose();
        // Add OLE object
        let oleObject = ws.OleObjects.Add(
          excelFileName,
          image,
          wasmModule.OleLinkType.Embed
        );
        oleObject.Location = ws.Range.get("B4");
        oleObject.ObjectType = wasmModule.OleObjectType.ExcelWorksheet;
        // Define the output file name
        const outputFileName = "InsertOLEObjects.xlsx";
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
