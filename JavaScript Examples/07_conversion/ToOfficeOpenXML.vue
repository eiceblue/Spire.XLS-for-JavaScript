<template>
  <span>Click the following button to convert an Excel file to an Open Office XML</span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {
        
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);
        
        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);
        // Set the text "Hello World" in cell A1 of the worksheet 
        sheet.Range.get("A1").Text = "Hello World";
        // Apply the color Gray25Percent to cell B1 using a known color
        sheet.Range.get("B1").Style.KnownColor = wasmModule.ExcelColors.Gray25Percent;
        // Apply the color Gold to cell C1 using a known color
        sheet.Range.get("C1").Style.KnownColor = wasmModule.ExcelColors.Gold;

        const outputFileName = 'ToOfficeOpenXML-out.xml';
        // Save the workbook as an XML file
        workbook.SaveAsXml({fileName:outputFileName});
        // Dispose of the workbook object to release resources
        workbook.Dispose();
        
        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/xml'});

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
