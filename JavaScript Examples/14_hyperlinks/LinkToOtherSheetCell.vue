<template>
  <span>Click the following button to add hyperlink to other sheet cell</span>
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
	
				//Get the first sheet
				const sheet = workbook.Worksheets.get(0);
	
				const range = sheet.Range.get("A1");
	
				//Add hyperlink in the range
				const hyperlink = sheet.HyperLinks.Add({range:range});
	
				//Set the link type
				hyperlink.Type = wasmModule.HyperLinkType.Workbook;
	
				//Set the display text
				hyperlink.TextToDisplay = "Link to Sheet2 cell C5";
	
				//Set the address
				hyperlink.Address = "Sheet2!C5";

        //Save result file
        const outputFileName = 'LinkToOtherSheetCell_out.xlsx';
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
