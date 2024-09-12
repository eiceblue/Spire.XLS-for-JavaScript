<template>
  <span>Click the following button to fit one page width when converting to PDF </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref} from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      if (wasmModule) {
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='FitWidthWhenConvertToPDF.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        for (let i = 0; i < workbook.Worksheets.Count; i++) {
            let sheet = workbook.Worksheets.get(i);
            // Auto fit page height
            sheet.PageSetup.FitToPagesTall = 0;
            // Fit one page width
            sheet.PageSetup.FitToPagesWide = 1;
        }

        const outputFileName = 'FitWidthWhenConvertToPDF-out.pdf';
        //Save to PDF document
        workbook.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PDF});
        // Dispose of the workbook object to release resources
        workbook.Dispose();
        
        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/pdf'});

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
