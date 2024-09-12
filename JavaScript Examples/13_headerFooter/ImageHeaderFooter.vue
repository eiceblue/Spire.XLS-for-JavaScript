<template>
  <span>Click the following button to insert image into Header or footer</span>
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
        let excelFileName='ImageHeaderFooter.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        let imgFileName='logo.png';
        await wasmModule.FetchFileToVFS(imgFileName, '', `${import.meta.env.BASE_URL}static/image/`);


        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        // Get the first sheet
        const sheet = workbook.Worksheets.get(0);

        // Load an image 
        const image = wasmModule.Stream.CreateByFile( imgFileName);

        // Set the image header
        sheet.PageSetup.LeftHeaderImage = image;
        sheet.PageSetup.LeftHeader = "&G";

        // Set the image footer
        sheet.PageSetup.CenterFooterImage = image;
        sheet.PageSetup.CenterFooter = "&G";

        // Set the view mode of the sheet
        sheet.ViewMode = wasmModule.ViewMode.Layout;

        //Save result file
        const outputFileName = 'ImageHeaderFooter_out.xlsx';
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
