<template>
  <span>Click the following button to set the order of shapes in the worksheet</span>
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
        let excelFileName = 'SetShapeOrder.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Bring the picture forward one level
        workbook.Worksheets.get(0).Pictures.get(0).ChangeLayer(wasmModule.ShapeLayerChangeType.BringForward);

        //Bring the image in fron of all other objects
        workbook.Worksheets.get(1).Pictures.get(0).ChangeLayer(wasmModule.ShapeLayerChangeType.BringToFront);

        //Send the shape back one level
        let shape = workbook.Worksheets.get(2).PrstGeomShapes.get(1);
        shape.ChangeLayer(wasmModule.ShapeLayerChangeType.SendBackward);

        //Send the shape behind all other objects
        shape = workbook.Worksheets.get(3).PrstGeomShapes.get(1);
        shape.ChangeLayer(wasmModule.ShapeLayerChangeType.SendToBack);

        // Save the modified workbook 
        const outputFile = 'SetShapeOrder.xlsx';
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
