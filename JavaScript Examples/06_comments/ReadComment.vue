<template>
  <span>Click the following button to read comment</span>
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
      if (wasmModule) {
        wasmModule = window.wasmModule;
        
        let inputFileName='ReadComment.xls';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        // Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        let builder = [];
        
        // Get the comment text
        builder.push(sheet.Range.get("A1").Comment.Text + "\n\t");
        builder.push(sheet.Range.get("A2").Comment.RichText.RtfText);
        
        const outputFileName = 'ReadComment-out.txt';
        // Save them to a txt file
        writeTextToFile(builder.join('\n'),outputFileName);
        workbook.Dispose();
        
        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'text/plain'});

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
function writeTextToFile(text, filename) {
    FS.writeFile(filename, text, (err) => {
        if (err) throw err;
        console.log('The file has been saved!');
    });
}
</script>
