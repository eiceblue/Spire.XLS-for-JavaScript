<template>
  <span>Click the following button to convert Excel to SVG</span>
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

        let inputFileName='ToSVG.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        let i=0;
        for(let worksheet of workbook.Worksheets) {
            const outputFileName = "sheet-"+i+".svg";
            // Create a FileStream to write the SVG content to a file
            let fs = wasmModule.Stream.CreateByFile(outputFileName);
            // Convert the worksheet to SVG and write it to the FileStream
            worksheet.ToSVGStream(fs,0, 0,0, 0);
            fs.Flush();
            fs.Dispose();

            const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
            const modifiedFile = new Blob([modifiedFileArray], { type: 'image/svg+xml'});
            downloadName.value = outputFileName;
            downloadUrl.value = URL.createObjectURL(modifiedFile);
            i++;
        }
        // Dispose of the object to release resources
        workbook.Dispose();
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
