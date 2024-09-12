<template>
  <span>Click the following button to add rectangle shape in Excel file</span>
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
        let excelFileName = 'ExcelSample_N1.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
       
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Add rectangle shape 1------Rect
        let rect1 = sheet.RectangleShapes.AddRectangle(11, 2, 60, 100, wasmModule.RectangleShapeType.Rect);
        rect1.Line.Weight = 1;
        // Fill shape with solid color
        rect1.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
        rect1.Fill.ForeColor = wasmModule.Color.get_DarkGreen();

        // Add rectangle shape 2------RoundRect
        let rect2 = sheet.RectangleShapes.AddRectangle(11, 5, 60, 100,wasmModule.RectangleShapeType.RoundRect);
        rect2.Line.Weight = 1;
        rect2.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
        rect2.Fill.ForeColor = wasmModule.Color.get_DarkCyan();
 
        // Save the modified workbook 
        const outputFile = 'AddRectangleShape.xlsx';
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
