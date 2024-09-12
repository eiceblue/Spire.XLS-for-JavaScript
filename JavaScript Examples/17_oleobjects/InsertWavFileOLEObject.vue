<template>
  <span>The example demonstrates how to insert a wav file as OLE object</span>
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
        let pngFileName = "SpireXls.png";
        await wasmModule.FetchFileToVFS(
          pngFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );
        let wavFileName = "WAVFileSample.wav";
        await wasmModule.FetchFileToVFS(
          wavFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook
        const book = wasmModule.Workbook.Create();
        // Get the first worksheet
        let sheet = book.Worksheets.get(0);

        // Add OLE object
        let fs = wasmModule.Stream.CreateByFile(pngFileName);
        let oleObject = sheet.OleObjects.Add(
          wavFileName,
          fs,
          wasmModule.OleLinkType.Embed
        );

        // Set the object location
        oleObject.Location = sheet.Range.get("B4");
        // Set the object type as package
        oleObject.ObjectType = wasmModule.OleObjectType.Package;
        // Define the output file name
        const outputFileName = "InsertWavFileOLEObject.xlsx";
        // Save the workbook to the specified path
        book.SaveToFile({
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
        book.Dispose();
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
