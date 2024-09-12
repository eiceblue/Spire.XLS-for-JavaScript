<template>
  <span>Click the following button to check whether the cell row height or column width is auto fit</span>
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
        let excelFileName='CheckAutoFitRowsAndColumns.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        const result = [];

        // Gets whether the cell has an adaptive row height set
        const isRowAutofit = workbook.Worksheets.get(0).GetRowIsAutoFit(2);
        if (isRowAutofit) {
          result.push("The second row is auto fit row height.");
        } else {
          result.push("The second row is not auto fit row height.");
        }

        // Gets whether the cell has an adaptive column width set
        const isColAutofit = workbook.Worksheets.get(0).GetColumnIsAutoFit(2);
        if (isColAutofit) {
          result.push("The second column is auto fit column width.");
        } else {
          result.push("The second column is not auto fit column width.");
        }
        
        // Save result file
        const outputFileName = 'CheckAutoFitRowOrColumn_out.txt';
        FS.writeFile(outputFileName, result.join("\n"))

        // Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'text/plain'});

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
