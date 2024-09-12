<template>
  <span
    >Click the following button to lock specific column in a new Excel
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

        //Create a workbook
        const workbook = wasmModule.Workbook.Create();
        // Create an empty worksheet
        workbook.CreateEmptySheet();

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Loop through all the columns in the worksheet and unlock them
        for (let i = 0; i < 20; i++) {
          sheet.Rows.get(i).Style.Locked = false;
        }

        // Lock the fourth column in the worksheet
        sheet.Columns.get(3).Text = "Locked";
        sheet.Columns.get(3).Style.Locked = true;

        // Set the password
        sheet.Protect("123", wasmModule.SheetProtectionType.All);

        let outputFileName = "LockSpecificColumnInNewExcel_output.xlsx";
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
  