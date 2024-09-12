<template>
  <span>Click the following button to add comment with author</span>
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

        let inputFileName='AddCommentWithAuthor.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Get the range that will add comment
        let range = sheet.Range.get("C1");

        //Set the author and comment content
        let author = "E-iceblue";
        let text = "This is demo to show how to add a comment with editable Author property.";

        //Add comment to the range and set properties
        let comment = range.AddComment();
        comment.Width = 200;
        comment.Visible = true;
        comment.Text = author + ":\n" + text;

        //Set the font of the author
        let font = workbook.CreateFont();
        font.FontName = "Arial";
        font.KnownColor = wasmModule.ExcelColors.Black;
        font.IsBold = true;
        comment.RichText.SetFont(0, author.length, font);
        
        const outputFileName = 'AddCommentWithAuthor-out.xlsx';
        // Save the modified workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
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
