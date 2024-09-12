<template>
  <span>Click the following button to add oval shape in Excel file</span>
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
        // Fetch the image and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS("logo.png", '', `${import.meta.env.BASE_URL}static/image/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Add oval shape1
        let ovalShape1 = sheet.OvalShapes.AddOval(11, 2, 100, 100);
        ovalShape1.Line.Weight = 0;
        // Fill shape with solid color
        ovalShape1.Fill.FillType = wasmModule.ShapeFillType.SolidColor;
        ovalShape1.Fill.ForeColor = wasmModule.Color.get_DarkCyan();

        // Add oval shape2
        let ovalShape2 = sheet.OvalShapes.AddOval(11, 5, 100, 100);
        ovalShape2.Line.Weight = 1;
        // Fill shape with picture
        ovalShape2.Line.DashStyle = wasmModule.ShapeDashLineStyleType.Solid;
   
        ovalShape2.Fill.CustomPicture("logo.png");

        // Save the modified workbook 
        const outputFile = 'AddOvalShape.xlsx';
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
