<template>
  <span>Click the following button to add hyperlink to text in excel workbook</span>
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
        let excelFileName='CommonTemplate1.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);

        //Add url link
        const UrlLink = sheet.HyperLinks.Add({range:sheet.Range.get("D10")});
        UrlLink.TextToDisplay = sheet.Range.get("D10").Text;
        UrlLink.Type = wasmModule.HyperLinkType.Url;
        UrlLink.Address = "http://en.wikipedia.org/wiki/Chicago";

        //Add email link
        const MailLink = sheet.HyperLinks.Add({range:sheet.Range.get("E10")});
        MailLink.TextToDisplay = sheet.Range.get("E10").Text;
        MailLink.Type = wasmModule.HyperLinkType.Url;
        MailLink.Address = "mailto:Amor.Aqua@gmail.com";

        //Save result file
        const outputFileName = 'AddHyperlinkToText_out.xlsx';
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
