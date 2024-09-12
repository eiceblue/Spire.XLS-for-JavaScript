<template>
  <span>Convert each worksheet to different PDF</span>
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

        let inputFileName='EachWorksheetToDifferentPDF.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});
        
        //Save each sheet to PDF
        for (let i = 0; i < workbook.Worksheets.Count; i++) {
            let sheet = workbook.Worksheets.get(i);    
            const outputFileName = sheet.Name + '.pdf';
            sheet.SaveToPdf({fileName: outputFileName});

            const modifiedFileArray = FS.readFile(outputFileName);
            const modifiedFile = new Blob([modifiedFileArray], {type: 'application/pdf'});

            downloadName.value = outputFileName;
            downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
      // Dispose of the workbook object to release resources
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
