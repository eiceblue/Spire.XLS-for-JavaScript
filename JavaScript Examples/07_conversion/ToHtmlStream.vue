<template>
<!--待修改部分：这里修改为每个功能对应的描述 -->
  <span>Click the following button to save a worksheet to an html stream  </span>
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

        let inputFileName='ToHtmlStream.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        
        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        //Set the html options
        let options = wasmModule.HTMLOptions.Create();
        options.ImageEmbedded = true;
        const outputFileName = 'ToHtmlStream-out.html';
        //Save sheet to html stream
        let fileStream = wasmModule.Stream.CreateByFile(outputFileName);
        sheet.SaveToHtml({stream:fileStream, saveOption:options});
        // Dispose of the object to release resources
        fileStream.Dispose();
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'text/html'});

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
