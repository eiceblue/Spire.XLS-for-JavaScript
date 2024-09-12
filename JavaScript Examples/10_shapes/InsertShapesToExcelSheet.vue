<template>
  <span>Click the following button to insert shapes to Excel worksheet</span>
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
        // Fetch the image and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS("spirexls.png", '', `${import.meta.env.BASE_URL}static/image/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 

        let sheet = workbook.Worksheets.get(0);

        //Add a triangle shape.
        let triangle = sheet.PrstGeomShapes.AddPrstGeomShape(2, 2, 100, 100, wasmModule.PrstGeomShapeType.Triangle);
        //Fill the triangle with solid color.
        triangle.Fill.ForeColor = wasmModule.Color.get_Yellow();
        triangle.Fill.FillType = wasmModule.ShapeFillType.SolidColor;

        //Add a heart shape.
        let heart = sheet.PrstGeomShapes.AddPrstGeomShape(2, 5, 100, 100, wasmModule.PrstGeomShapeType.Heart);
        //Fill the heart with gradient color.
        heart.Fill.ForeColor = wasmModule.Color.get_Red();
        heart.Fill.FillType = wasmModule.ShapeFillType.Gradient;

        //Add an arrow shape with default color.
        let arrow = sheet.PrstGeomShapes.AddPrstGeomShape(10, 2, 100, 100, wasmModule.PrstGeomShapeType.CurvedRightArrow);

        //Add a cloud shape.
        let cloud = sheet.PrstGeomShapes.AddPrstGeomShape(10, 5, 100, 100, wasmModule.PrstGeomShapeType.Cloud);
        //Fill the cloud with custom picture
        cloud.Fill.CustomPicture({im:wasmModule.Stream.CreateByFile("wasmModule.png"), name:"wasmModule.png"});

        cloud.Fill.FillType = wasmModule.ShapeFillType.Picture;

        // Save the modified workbook 
        const outputFile = 'InsertShapesToExcelSheet.xlsx';
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
