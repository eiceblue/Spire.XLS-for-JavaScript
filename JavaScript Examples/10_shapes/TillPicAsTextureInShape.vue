<template>
  <span>Click the following button to till picture as texture in a shape</span>
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
        let excelFileName = 'TillPicAsTextureInShape.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        // Fetch the image and add it to the Virtual File System (VFS)
        await wasmModule.FetchFileToVFS("Logo.png", '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Get the first shape
        let shape = sheet.PrstGeomShapes.get(0);

        //Fill shape with texture
        shape.Fill.FillType = wasmModule.ShapeFillType.Texture;

        //Custom texture with picture
        shape.Fill.CustomTexture({path:"Logo.png"});

        //Tile pciture as texture
        shape.Fill.Tile = true;
      
        // Save the modified workbook   
        const outputFile = 'TillPicAsTextureInShape.xlsx';
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
