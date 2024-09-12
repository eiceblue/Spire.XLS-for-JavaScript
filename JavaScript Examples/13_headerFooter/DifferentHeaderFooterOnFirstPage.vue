<template>
  <span>Click the following button to add different header and footer for the first page in Excel file</span>
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

        // Create a workbook
        const workbook = wasmModule.Workbook.Create();

        // Get the first worksheet
        const sheet = workbook.Worksheets.get(0);
        sheet.Range.get("A1").Text = "Hello World";
        sheet.Range.get("F30").Text = "Hello World";
        sheet.Range.get("G150").Text = "Hello World";

        // Set the value to show the headers/footers for first page are different from the other pages
        sheet.PageSetup.DifferentFirst = 1;

        // Set the header and footer for the first page
        sheet.PageSetup.FirstHeaderString = "Different First page";
        sheet.PageSetup.FirstFooterString = "Different First footer";

        // Set the other pages' header and footer
        sheet.PageSetup.LeftHeader = "Demo of Spire.XLS";
        sheet.PageSetup.CenterFooter = "Footer by Spire.XLS";

        //Save result file
        const outputFileName = 'DifferentHeaderFooterOnFirstPage_out.xlsx';
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
