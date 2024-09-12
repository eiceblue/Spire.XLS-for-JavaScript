<template>
  <span>Click the following button to set position and alignment for comments</span>
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

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        //Set two font styles which will be used in comments
        let font1 = workbook.CreateFont();
        font1.FontName = "Calibri";
        font1.Color = wasmModule.Color.get_Firebrick();
        font1.IsBold = true;
        font1.Size = 12;
        let font2 = workbook.CreateFont();
        font2.FontName = "Calibri";
        font2.Color = wasmModule.Color.get_Blue();
        font2.Size = 12;
        font2.IsBold = true;

        //Add comment 1 and set its size, text, position and alignment
        sheet.Range.get("G5").Text = "Spire.XLS";
        let Comment1 = sheet.Range.get("G5").Comment;
        Comment1.IsVisible = true;
        Comment1.Height = 150;
        Comment1.Width = 300;
        Comment1.RichText.Text = "Spire.XLS for JavaScript:\nStandalone Excel component to meet your needs for conversion, data manipulation, charts in workbook etc. ";
        Comment1.RichText.SetFont(0, 19, font1);
        Comment1.TextRotation = wasmModule.TextRotationType.LeftToRight;

        //Set the position of Comment
        Comment1.Top = 20;
        Comment1.Left = 40;

        //Set the alignment of text in Comment
        Comment1.VAlignment = wasmModule.CommentVAlignType.Center;
        Comment1.HAlignment = wasmModule.CommentHAlignType.Justified;

        //Add comment2 and set its size, text, position and alignment for comparison
        sheet.Range.get("D14").Text = "E-iceblue";
        let Comment2 = sheet.Range.get("D14").Comment;
        Comment2.IsVisible = true;
        Comment2.Height = 150;
        Comment2.Width = 300;
        Comment2.RichText.Text = "About E-iceblue: \nWe focus on providing excellent office components for developers to operate Word, Excel, PDF, and PowerPoint documents.";
        Comment2.TextRotation = wasmModule.TextRotationType.LeftToRight;
        Comment2.RichText.SetFont(0, 16, font2);
        //Set the position of Comment
        Comment2.Top = 170;
        Comment2.Left = 450;
        //Set the alignment of text in Comment
        Comment2.VAlignment = wasmModule.CommentVAlignType.Top;
        Comment2.HAlignment = wasmModule.CommentHAlignType.Justified;
        
        const outputFileName = 'SetPositionAndAlignment-out.xlsx';
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
