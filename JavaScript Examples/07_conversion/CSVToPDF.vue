<template>
  <span>Click the following button to convert CSV to PDF</span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {

        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        let inputFileName='CSVToPDF.csv';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing CSV document
        workbook.LoadFromFile({ fileName: inputFileName, separator: ",", row: 1, column: 1 });

        // Set the SheetFitToPage property as true
        workbook.ConverterSetting.SheetFitToPage = true;

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Autofit a column if the characters in the column exceed column width
        for (let i = 1; i < sheet.Columns.Count; i++) {
            sheet.AutoFitColumn(i);
        }

        const outputFileName = 'CSVToPDF-out.pdf';
        //Save to PDF document
        workbook.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});
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
