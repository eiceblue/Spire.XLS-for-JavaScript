<template>
  <span>Click the following button to insert comment</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {

        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);
        
        let inputFileName='WriteComment.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        
        // Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        //Creates font
        let font = workbook.CreateFont();
        font.FontName = "Arial";
        font.Size = 11;
        font.KnownColor = wasmModule.ExcelColors.Orange;
        let fontBlue = workbook.CreateFont();
        fontBlue.KnownColor = wasmModule.ExcelColors.LightBlue;
        let fontGreen = workbook.CreateFont();
        fontGreen.KnownColor = wasmModule.ExcelColors.LightGreen;

        let range = sheet.Range.get("B11");
        range.Text = "Regular comment";
        range.Comment.Text = "Regular comment";
        range.AutoFitColumns();

        range = sheet.Range.get("B12");
        range.Text = "Rich text comment";
        range.RichText.SetFont(0, 16, font);
        range.AutoFitColumns();
        //Rich text comment
        range.Comment.RichText.Text = "Rich text comment";
        range.Comment.RichText.SetFont(0, 4, fontGreen);
        range.Comment.RichText.SetFont(5, 9, fontBlue);
        
        const outputFileName = 'WriteComment-out.xlsx';
        // Save the modified workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
