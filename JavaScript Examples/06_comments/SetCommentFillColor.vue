<template>
  <span>Click the following button to set the fill color of a comment</span>
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

        //Create Excel font
        let font = workbook.CreateFont();
        font.FontName = "Arial";
        font.Size = 11;
        font.KnownColor = wasmModule.ExcelColors.Orange;

        //Add the comment
        let range = sheet.Range.get("A1");
        let commentText = "This is a comment";
        range.Comment.Text = commentText;
        range.Comment.RichText.SetFont(0, commentText.length - 1, font);

        //Set comment Color
        range.Comment.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
        range.Comment.Fill.ForeColor = wasmModule.Color.get_SkyBlue();
        range.Comment.Visible = true;
        
        const outputFileName = 'SetCommentFillColor-out.xlsx';
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
