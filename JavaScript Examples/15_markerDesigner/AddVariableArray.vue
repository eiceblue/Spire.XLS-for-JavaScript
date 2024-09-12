<template>
  <span>Click the following button to add variable array to excel workbook</span>
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
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

				//Create a workbook
				const workbook = wasmModule.Workbook.Create();
	
				//Get the first worksheet
				const sheet = workbook.Worksheets.get(0);
	
				//Set marker designer field in cell A1
				sheet.Range.get("A1").Value = "&=Array";
	
				//Fill Array
				workbook.MarkerDesigner.AddArray("Array",
					[wasmModule.String.Create("Spire.Xls"),
					wasmModule.String.Create("Spire.Doc"),
					wasmModule.String.Create("Spire.PDF"),
					wasmModule.String.Create("Spire.Presentation"),
					wasmModule.String.Create("Spire.Email")]);
				workbook.MarkerDesigner.Apply();
				workbook.CalculateAllValue();
	
				//AutoFit
				sheet.AllocatedRange.AutoFitRows();
				sheet.AllocatedRange.AutoFitColumns();

        //Save result file
        const outputFileName = 'AddVariableArray_out.xlsx';
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
