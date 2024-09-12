<template>
  <span>Click the following button to convert CSV to Excel </span>
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

        let inputFileName='CSVToExcel.csv';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
                
        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
         // Load an existing CSV document
        workbook.LoadFromFile({ fileName: inputFileName, separator: ",", row: 1, column: 1 });
        // Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);
        // Ignore error options for the range D2:E19, treating numbers as text
        sheet.Range.get("D2:E19").IgnoreErrorOptions = wasmModule.IgnoreErrorType.NumberAsText;
        // Auto-fit columns in the allocated range of the worksheet
        sheet.AllocatedRange.AutoFitColumns();

        const outputFileName = 'CSVToExcel-out.xlsx';
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
