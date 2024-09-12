<template>
  <span>Click the following button to write hyperlink</span>
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
        let excelFileName='WriteHyperlinks.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
				const sheet = workbook.Worksheets.get(0);
	
				//Set links
				sheet.Range.get("B9").Text = "Home page";
				const hylink1 = sheet.HyperLinks.Add({range:sheet.Range.get("B10")});
				hylink1.Type = wasmModule.HyperLinkType.Url;
				hylink1.Address = "http://www.e-iceblue.com";
	
				sheet.Range.get("B11").Text = "Support";
				const hylink2 = sheet.HyperLinks.Add({range:sheet.Range.get("B12")});
				hylink2.Type = wasmModule.HyperLinkType.Url;
				hylink2.Address = "mailto:support@e-iceblue.com";
	
				sheet.Range.get("B13").Text = "Forum";
				const hylink3 = sheet.HyperLinks.Add({range:sheet.Range.get("B14")});
				hylink3.Type = wasmModule.HyperLinkType.Url;
				hylink3.Address = "https://www.e-iceblue.com/forum/";

        //Save result file
        const outputFileName = 'WriteHyperlinks_out.xlsx';
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
