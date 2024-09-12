<template>
  <span>Click the following button to add line shapes in Excel file</span>
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

        // Add shape line1
        let line1 = sheet.Lines.AddLine({ row: 10, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.Line });
        line1.DashStyle = wasmModule.ShapeDashLineStyleType.Solid;
        line1.Color = wasmModule.Color.get_CadetBlue();
        line1.Weight = 2;
        line1.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

        // Add shape line2
        let line2 = sheet.Lines.AddLine({ row: 12, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.CurveLine });
        line2.DashStyle = wasmModule.ShapeDashLineStyleType.Dotted;
        line2.Color = wasmModule.Color.get_OrangeRed();
        line2.Weight = 2;

        // Add shape line3
        let line3 = sheet.Lines.AddLine({ row: 14, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.ElbowLine });
        line3.DashStyle = wasmModule.ShapeDashLineStyleType.DashDotDot;
        line3.Color = wasmModule.Color.get_Purple();
        line3.Weight = 2;

        // Add shape line4
        let line4 = sheet.Lines.AddLine({ row: 16, column: 2, width: 200, height: 1, lineShapeType: wasmModule.LineShapeType.LineInv });
        line4.DashStyle = wasmModule.ShapeDashLineStyleType.Dashed;
        line4.Color = wasmModule.Color.get_Green();
        line4.Weight = 2;

        // Save the modified workbook 
        const outputFile = 'AddLineShape.xlsx';
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
