<template>
  <span>Click the following button to draw one line through two points in Excel</span>
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
        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        let worksheet = workbook.Worksheets.get(0);

        //1)Draw a line according to relative position
        let line1 = worksheet.TypedLines.AddLine();
        line1.LeftColumn = 3;
        line1.TopRow = 3;
        line1.LeftColumnOffset = 0;
        line1.TopRowOffset = 0;

        line1.RightColumn = 4;
        line1.BottomRow = 5;
        line1.RightColumnOffset = 0;
        line1.BottomRowOffset = 0;

        //2)Draw a line according to absolute position(pixels).
        let line2 = worksheet.TypedLines.AddLine();
        line2.StartPoint = wasmModule.Point.Create(30, 50);
        line2.EndPoint = wasmModule.Point.Create(20, 80);
      
        // Save the modified workbook   
        const outputFile = 'DrawOneLineThroughTwoPoints.xlsx';
        workbook.SaveToFile(outputFile);
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved Excel file from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Download the converted Excel file
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
