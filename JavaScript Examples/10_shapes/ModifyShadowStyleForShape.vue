<template>
  <span>Click the following button to set shadow style for shape in an existing excel file</span>
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
        let excelFileName = 'Template_Xls_5.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the first worksheet.
        let sheet = workbook.Worksheets.get(0);

        //Get the third shape from the worksheet.
        let shape = sheet.PrstGeomShapes.get(2);

        //Set the shadow style for the shape.
        shape.Shadow.Angle = 90;
        shape.Shadow.Transparency = 30;
        shape.Shadow.Distance = 10;
        shape.Shadow.Size = 130;
        shape.Shadow.Color = wasmModule.Color.get_Yellow();
        shape.Shadow.Blur = 30;
        shape.Shadow.HasCustomStyle = true;

        // Save the modified workbook 
        const outputFile = 'ModifyShadowStyleForShape.xlsx';
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
