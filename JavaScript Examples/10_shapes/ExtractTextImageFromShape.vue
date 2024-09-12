<template>
  <span>Click the following button to extract Text and image from shape </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated txt 
  </a>
  <a v-if="downloadUrl1" :href="downloadUrl1" :download="downloadName1">
    Click here to download the generated image
  </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");
    const downloadUrl1 = ref(null);
    const downloadName1 = ref("");
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
        //Extract text from the first shape and save to a txt file.
        let shape1 = sheet.PrstGeomShapes.get(2);
        let s = shape1.Text;
        let sb = [];
        sb.push(`The text in the third shape is: ${s}`);
        
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Convert the text to a Blob
        const textFile = 'ExtractTextImageFromShape.txt';
        const modifiedFile = new Blob([sb.toString], { type: "text/plain;charset=utf-8"});

        // Download the txt file
        downloadName.value = textFile;
        downloadUrl.value = URL.createObjectURL(modifiedFile);

        let shape2 = sheet.PrstGeomShapes.get(1);

        let image = shape2.Fill.Picture;
        let imageFile = `ExtractTextImageFromShape.png`;
        image.Save(imageFile);
        // Read the saved image from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(imageFile);
        const modifiedFile1 = new Blob([modifiedFileArray], { type: 'application/png' });

        // Download the image
        downloadName1.value = imageFile;
        downloadUrl1.value = URL.createObjectURL(modifiedFile1);

      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
      downloadName1,
      downloadUrl1
    };
  }
};
</script>
