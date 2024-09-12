<template>
  <span>Click the following button to copy rows</span>
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
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Input file
        let excelFileName='Copying.xls';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first and second sheet
        let sheet1 = workbook.Worksheets.get(1);
        let sheet2 = workbook.Worksheets.get(0);
        
        //Copy the first row to the third row in the same sheet
        sheet1.Copy({sourceRange:sheet1.Rows.get(0), destRange:sheet1.Rows.get(2), copyStyle:true, updateReference:true, ignoreSize:true});
        
        //Copy the first row to the second row in the different sheet
        sheet1.Copy({sourceRange:sheet1.Rows.get(0), destRange:sheet2.Rows.get(1), copyStyle:true, updateReference:true, ignoreSize:true});

        //Save result file
        const outputFileName = 'CopyRows_out.xlsx';
        workbook.SaveToFile({fileName: outputFileName, version:wasmModule.ExcelVersion.Version2010});

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        // Download the result file
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
