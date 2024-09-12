<template>
  <span>Click the following button to add comment with picture</span>
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
        let inputFileName='logo.png';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);
        // Set value for the range
        sheet.Range.get("C6").Text = "E-iceblue";
        //Add the comment
        let comment = sheet.Range.get("C6").AddComment();
        // Fill the comment with a customized background picture
        comment.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile(inputFileName), name:"None"});
        comment.Visible = true;
        
        const outputFileName = 'AddCommentWithPicture-out.xlsx';
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
