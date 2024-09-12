<template>
  <span>Click the following button to add arrow lines to Excel file</span>
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
        
        // Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        // Add a Double Arrow and fill the line with solid color.
        let line = sheet.TypedLines.AddLine();
        line.Top = 10;
        line.Left = 20;
        line.Width = 100;
        line.Height = 0;
        line.Color = wasmModule.Color.get_Blue();
        line.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
        line.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

        // Add an Arrow and fill the line with solid color.
        let line_1 = sheet.TypedLines.AddLine();
        line_1.Top = 50;
        line_1.Left = 30;
        line_1.Width = 100;
        line_1.Height = 100;
        line_1.Color = wasmModule.Color.get_Red();
        line_1.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineNoArrow;
        line_1.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

        // Add an Elbow Arrow Connector.
        let line3 = sheet.TypedLines.AddLine();
        line3.LineShapeType = wasmModule.LineShapeType.ElbowLine;
        line3.Width = 30;
        line3.Height = 50;
        line3.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
        line3.Top = 100;
        line3.Left = 50;

        // Add an Elbow Double-Arrow Connector.
        let line2 = sheet.TypedLines.AddLine();
        line2.LineShapeType = wasmModule.LineShapeType.ElbowLine;
        line2.Width = 50;
        line2.Height = 50;
        line2.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
        line2.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;
        line2.Left = 120;
        line2.Top = 100;

        // Add a Curved Arrow Connector.
        line3 = sheet.TypedLines.AddLine();
        line3.LineShapeType = wasmModule.LineShapeType.CurveLine;
        line3.Width = 30;
        line3.Height = 50;
        line3.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowOpen;
        line3.Top = 100;
        line3.Left = 200;

        // Add a Curved Double-Arrow Connector.
        line2 = sheet.TypedLines.AddLine();
        line2.LineShapeType = wasmModule.LineShapeType.CurveLine;
        line2.Width = 30;
        line2.Height = 50;
        line2.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowOpen;
        line2.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowOpen;
        line2.Left = 250;
        line2.Top = 100;

        // Save the modified workbook 
        const outputFile = 'AddArrowLineToExcelFile.xlsx';
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
