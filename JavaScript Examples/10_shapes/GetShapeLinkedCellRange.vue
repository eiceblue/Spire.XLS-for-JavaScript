<template>
  <span>Click the following button to get shape linked cell range address</span>
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
        let excelFileName = 'CellLinkedRangeLocal.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
      
        let sb = [];
        workbook.LoadFromFile(excelFileName);
        let sheet = workbook.Worksheets.get(0);
        let prstGeomShapeCollection = sheet.PrstGeomShapes;
        let shape = prstGeomShapeCollection.get({name:"Yesterday"});
        let cellAddress = shape.LinkedCell.RangeAddress;
        sb.push(`${cellAddress}\n`);
        shape = prstGeomShapeCollection.get({name:"NewShapes"});
        cellAddress = shape.LinkedCell.RangeAddress;
        sb.push(cellAddress);

        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Create a Blob object 
        const outputFile = 'GetShapeLinkedCellRange.txt';
        const modifiedFile = new Blob([sb.toString()], { type: "text/plain;charset=utf-8"});

        // Download the txt file
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
