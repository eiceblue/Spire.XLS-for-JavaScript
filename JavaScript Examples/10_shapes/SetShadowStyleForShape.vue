<template>
  <span>Click the following button to set shape shadow style for newly Excel file</span>
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

        //Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        //Add an ellipse shape.
        let ellipse = sheet.PrstGeomShapes.AddPrstGeomShape(5, 5, 150, 100, wasmModule.PrstGeomShapeType.Ellipse);

        //Set the shadow style for the ellipse.
        ellipse.Shadow.Angle = 90;
        ellipse.Shadow.Distance = 10;
        ellipse.Shadow.Size = 150;
        ellipse.Shadow.Color = wasmModule.Color.get_Gray();
        ellipse.Shadow.Blur = 30;
        ellipse.Shadow.Transparency = 1;
        ellipse.Shadow.HasCustomStyle = true;

        // Save the modified workbook 
        const outputFile = 'SetShadowStyleForShape.xlsx';
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
