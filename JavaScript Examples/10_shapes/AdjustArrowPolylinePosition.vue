<template>
  <span>Click the following button to adjust the position of arrow polyline in Excel</span>
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

        // Draw an elbow arrow
        let line = worksheet.TypedLines.AddLine({
            row: 5,
            column: 5,
            width: 100,
            height: 100,
            lineShapeType: wasmModule.LineShapeType.ElbowLine
        });
        line.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineNoArrow;
        line.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
        let ad = line.ShapeAdjustValues.AddAdjustValue(wasmModule.GeomertyAdjustValueFormulaType.LiteralValue);

        // When the parameter value is less than 0, the focus of the line is on the left side of the left point, when it is equal to 0, the position is the same as the left point, it is equal to 50 in the middle of the graph, and when it is equal to 100, it is the same as the right point.
        ad.SetFormulaParameter([-50]);

        // Save the modified workbook 
        const outputFile = 'AdjustArrowPolylinePosition.xlsx';
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
