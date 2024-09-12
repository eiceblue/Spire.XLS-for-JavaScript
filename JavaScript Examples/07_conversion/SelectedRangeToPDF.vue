<template>
  <span>Click the following button to convert selected range of cells to PDF </span>
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

        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='SelectedRangeToPDF.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        // Add a new sheet to workbook
        workbook.Worksheets.Add("newsheet");
        // Copy your area to new sheet.
        workbook.Worksheets.get(0).Range.get("A9:E15").Copy({destRange:workbook.Worksheets.get(1).Range.get("A9:E15"), updateReference:false, copyStyles:true});
        // Auto fit column width
        workbook.Worksheets.get(1).Range.get("A9:E15").AutoFitColumns();

        const outputFileName = 'SelectedRangeToPDF-out.pdf';
        // Save worksheet to PDF
        workbook.Worksheets.get(1).SaveToPdf({fileName:outputFileName});
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
