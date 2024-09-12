<template>
  <span>Click the following button to copy sheet to another Excel file</span>
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

        //Create a workbook.
        let workbook = wasmModule.Workbook.Create();

        //Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        //Put some data into header rows (A1:A4)
        for (let i = 1; i <= 5; i++) {
          sheet.Range.get(`A${i}`).Text = `Header Row ${i}`;
        }

        //Put some detail data (A5:A99)
        for (let i = 5; i < 100; i++) {
          sheet.Range.get(`A${i}`).Text = `Detail Row ${i}`;
        }
        //Define a pagesetup object based on the first worksheet.
        let pageSetup = sheet.PageSetup;
        //The first five rows are repeated in each page. It can be seen in print preview.
        pageSetup.PrintTitleRows = "$1:$5";
        //Create another Workbook.
        let workbook1 = wasmModule.Workbook.Create();
        //Get the first worksheet in the book.
        let sheet1 = workbook1.Worksheets.get(0);
        //Copy worksheet to destination worsheet in another Excel file.
        sheet1.CopyFrom(sheet);

        const outputFileName1 = "CopySheetToAnotherXlsFile_output1.xlsx";
        const outputFileName2 = "CopySheetToAnotherXlsFile_output2.xlsx";
        workbook.SaveToFile({ fileName: outputFileName1 });
        workbook1.SaveToFile({ fileName: outputFileName2 });

        // Dispose of the workbook object to release resources
        workbook.Dispose();
        workbook1.Dispose();


        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName1);
        const modifiedFileArray2 = wasmModule.FS.readFile(outputFileName2);
        const modifiedFile = new Blob([modifiedFileArray, modifiedFileArray2], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

        // download the file
        downloadName.value = outputFileName1;
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
  