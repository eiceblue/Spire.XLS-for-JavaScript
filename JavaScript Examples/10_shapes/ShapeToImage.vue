<template>
  <span>Click the following button to convert shape to image </span>
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
        let excelFileName = 'ShapeToImage.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the first worksheet
        let sheet1 = workbook.Worksheets.get(0);

        //Get the first shape from the first worksheet
        let shape = sheet1.PrstGeomShapes.get(0);

        //Save the shape to a image
        let img = shape.SaveToImage();
        let outputFile="ShapeToImage.png"
        img.Save(outputFile);

        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved image from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/png' });

        // Download the image
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
