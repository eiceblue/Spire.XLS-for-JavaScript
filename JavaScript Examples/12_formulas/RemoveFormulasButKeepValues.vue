<template>
  <span>Click the following button to remove formula but keep its value in Excel file</span>
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
        let excelFileName='RemoveFormulasButKeepValues.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Loop through worksheets
        for (let i = 0; i < workbook.Worksheets.Count; i++) {
          let sheet = workbook.Worksheets.get(i);
          //Loop through cells
          for (const cell of sheet.Range.Cells) {
            //If the cell contains formula, get the formula value, clear cell content, and then fill the formula value into the cell.
            if (cell.HasFormula) {
              const value = cell.FormulaValue;
              cell.Clear(wasmModule.ExcelClearOptions.ClearContent);
              cell.Value2 = wasmModule.String.Create(value);
            }
          }
        }

        //Save result file
        const outputFileName = 'RemoveFormulasButKeepValues_out.xlsx';
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
