<template>
  <span>Click the following button to retrieve external file hyperlinks in Excel file</span>
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

        // Input file
        let excelFileName='RetrieveExternalFileHyperlinks.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
				const sheet = workbook.Worksheets.get(0);
	
				const content = [];
	
				//Retrieve external file hyperlinks.
				const hyperlinks = sheet.HyperLinks;
				for(let i=0; i<hyperlinks.Count; i++) {
					const item = hyperlinks.get(i);
					const address = item.Address;
					const sheetName = item.Range.WorksheetName;
					const range = item.Range;
					content.push(`Cell[${range.Row},${range.Column}] in sheet "${sheetName}" contains File URL: ${address}`);
				}

        //Save result file
        const outputFileName = 'RetrieveExternalFileHyperlinks_out.txt';
        FS.writeFile(outputFileName, content.join("\n"));

        //Dispose
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
