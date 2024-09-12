<template>
  <span
    >The example demonstrates how to extract ole objects from Excel file</span
  >
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrlDoc" :href="downloadUrlDoc" :download="downloadNameDoc"
    >Click here to download the generated Doc file</a
  >
  <a v-if="downloadUrlPdf" :href="downloadUrlPdf" :download="downloadNamePdf"
    >Click here to download the generated Pdf file</a
  >
  <a v-if="downloadUrlPpt" :href="downloadUrlPpt" :download="downloadNamePpt"
    >Click here to download the generated Ppt file</a
  >
</template>
<script>
import { ref } from "vue";
export default {
  setup() {
    const downloadUrlDoc = ref(null);
    const downloadNameDoc = ref("");
    const downloadUrlPdf = ref(null);
    const downloadNamePdf = ref("");
    const downloadUrlPpt = ref(null);
    const downloadNamePpt = ref("");

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
        let excelFileName = "ExtractOle2.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );
        // Create a new workbook
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile({
          fileName: excelFileName,
          version: wasmModule.ExcelVersion.Version2010,
        });
        // Get the first worksheet
        let sheet = book.Worksheets.get(0);
        const topostMessageDoc = (outputFileName, OleData) => {
          wasmModule.FS.writeFile(outputFileName, OleData);

          // Read the saved file and convert to a Blob object
          const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
          const modifiedFile = new Blob([modifiedFileArray], {
            type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          });
          // Download the Doc file

          downloadNameDoc.value = outputFileName;
          downloadUrlDoc.value = URL.createObjectURL(modifiedFile);
        };
        const topostMessagePdf = (outputFileName, OleData) => {
          wasmModule.FS.writeFile(outputFileName, OleData);

          // Read the saved file and convert to a Blob object
          const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
          const modifiedFile = new Blob([modifiedFileArray], {
            type: "application/pdf",
          });
          // Download the Pdf file

          downloadNamePdf.value = outputFileName;
          downloadUrlPdf.value = URL.createObjectURL(modifiedFile);
        };
        const topostMessagePpt = (outputFileName, OleData) => {
          wasmModule.FS.writeFile(outputFileName, OleData);

          // Read the saved file and convert to a Blob object
          const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
          const modifiedFile = new Blob([modifiedFileArray], {
            type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          });
          // Download the Ppt file

          downloadNamePpt.value = outputFileName;
          downloadUrlPpt.value = URL.createObjectURL(modifiedFile);
        };
        // Extract ole objects
        if (sheet.HasOleObjects) {
          for (let obj of sheet.OleObjects) {
            let type = obj.ObjectType;
            // Word document
            if (type === wasmModule.OleObjectType.WordDocument) {
              // Define the output file name
              const outputFileName = "ExtractOLEObjects.docx";
              topostMessageDoc(outputFileName, obj.OleData);
            }
            // Pdf document
            if (type === wasmModule.OleObjectType.AdobeAcrobatDocument) {
              // Define the output file name
              const outputFileName = "ExtractOLEObjects.pdf";
              topostMessagePdf(outputFileName, obj.OleData);
            }
            // Ppt document
            if (type === wasmModule.OleObjectType.PowerPointSlide) {
              // Define the output file name
              const outputFileName = "ExtractOLEObjects.pptx";
              topostMessagePpt(outputFileName, obj.OleData);
            }
          }
        }
        // Clean up resources
        book.Dispose();
      }
    };

    return {
      startProcessing,
      downloadUrlDoc,
      downloadNameDoc,
      downloadUrlPdf,
      downloadNamePdf,
      downloadUrlPpt,
      downloadNamePpt,
    };
  },
};
</script>
