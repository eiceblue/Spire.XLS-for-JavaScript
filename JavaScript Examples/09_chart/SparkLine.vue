<template>
  <span>Click the following button to create Spark Line </span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Fetch the Excel file and add it to the Virtual File System (VFS)
        let excelFileName = 'SparkLine.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first sheet
        let sheet = workbook.Worksheets.get(0);

        // Add sparkline
        let sparklineGroup = sheet.SparklineGroups.AddGroup({sparklineType:wasmModule.SparklineType.Line});
        let sparklines = sparklineGroup.Add();
        sparklines.Add({dataRange:sheet.Range.get("A2:D2"), referenceRange:sheet.Range.get("E2")});
        sparklines.Add({dataRange:sheet.Range.get("A3:D3"), referenceRange:sheet.Range.get("E3")});
        sparklines.Add({dataRange:sheet.Range.get("A4:D4"), referenceRange:sheet.Range.get("E4")});
        sparklines.Add({dataRange:sheet.Range.get("A5:D5"), referenceRange:sheet.Range.get("E5")});
        sparklines.Add({dataRange:sheet.Range.get("A6:D6"), referenceRange:sheet.Range.get("E6")});
        sparklines.Add({dataRange:sheet.Range.get("A7:D7"), referenceRange:sheet.Range.get("E7")});
        sparklines.Add({dataRange:sheet.Range.get("A8:D8"), referenceRange:sheet.Range.get("E8")});
        sparklines.Add({dataRange:sheet.Range.get("A9:D9"), referenceRange:sheet.Range.get("E9")});
        sparklines.Add({dataRange:sheet.Range.get("A10:D10"), referenceRange:sheet.Range.get("E10")});
        sparklines.Add({dataRange:sheet.Range.get("A11:D11"), referenceRange:sheet.Range.get("E11")});
        sparklines.Add({dataRange:sheet.Range.get("A2:D2"), referenceRange:sheet.Range.get("E2")});
        sparklines.Add({dataRange:sheet.Range.get("A2:D2"), referenceRange:sheet.Range.get("E2")});

        // Save the modified workbook
        const outputFile = 'SparkLine.xlsx'; 
        workbook.SaveToFile(outputFile);
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved Excel file from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Download the Excel file
        downloadName.value = outputFile;
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
