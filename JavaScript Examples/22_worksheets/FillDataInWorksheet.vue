<template>
  <span>Click the following button to fill data in worksheet</span>
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

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();

        //Get first worksheet of the workbook
        let worksheet = workbook.Worksheets.get(0);

        //Fill data
        worksheet.Range.get("A1").Style.Font.IsBold = true;
        worksheet.Range.get("B1").Style.Font.IsBold = true;
        worksheet.Range.get("C1").Style.Font.IsBold = true;
        worksheet.Range.get("A1").Text = "Month";
        worksheet.Range.get("A2").Text = "January";
        worksheet.Range.get("A3").Text = "February";
        worksheet.Range.get("A4").Text = "March";
        worksheet.Range.get("A5").Text = "April";
        worksheet.Range.get("B1").Text = "Payments";
        worksheet.Range.get("B2").NumberValue = 251;
        worksheet.Range.get("B3").NumberValue = 515;
        worksheet.Range.get("B4").NumberValue = 454;
        worksheet.Range.get("B5").NumberValue = 874;
        worksheet.Range.get("C1").Text = "Sample";
        worksheet.Range.get("C2").Text = "Sample1";
        worksheet.Range.get("C3").Text = "Sample2";
        worksheet.Range.get("C4").Text = "Sample3";
        worksheet.Range.get("C5").Text = "Sample4";

        //Set width for the second column
        worksheet.SetColumnWidth(2, 10);

        const outputFileName = "FillDataInWorksheet_output.xlsx";
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
  