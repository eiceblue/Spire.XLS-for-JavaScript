<template>
  <span>Click the following button to insert a background image</span>
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
        let inputFileName1 = "Background.png";
        await wasmModule.FetchFileToVFS(inputFileName1,"",`${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "Template_Xls_1.xlsx";
        await wasmModule.FetchFileToVFS(inputFileName2,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile(inputFileName2);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Open an image
        let bm = wasmModule.Stream.CreateByFile(inputFileName1);

        // Set the image to be background image of the worksheet
        sheet.PageSetup.BackgoundImage = bm;

        const outputFileName = "InsertExcelBackgroundImage-out.xlsx";
        // Save the modified workbook to the specified file
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2013,
        });
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

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