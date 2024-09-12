<template>
  <span>Click the following button to add image hyperlink in excel workbook</span>
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

        //Load image
        let imgFileName='logo.png';
        await wasmModule.FetchFileToVFS(imgFileName, '', `${import.meta.env.BASE_URL}static/image/`);

        //Create a workbook
        const workbook = wasmModule.Workbook.Create();

        const sheet = workbook.Worksheets.get(0);

        //Add the description text
        sheet.Columns.get(0).ColumnWidth = 22;
        sheet.Range.get("A1").Text = "Image Hyperlink";
        sheet.Range.get("A1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Top;

        //Insert an image to a specific cell
        let picture = sheet.Pictures.Add({topRow:2, leftColumn:1, fileName:imgFileName});
        //Add a hyperlink to the image
        picture.SetHyperLink("https://www.e-iceblue.com/Misc/about-us.html", true);

        //Save result file
        const outputFileName = 'AddImageHyperlink_out.xlsx';
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
