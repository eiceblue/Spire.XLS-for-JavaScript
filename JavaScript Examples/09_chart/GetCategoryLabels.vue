<template>
  <span>Click the following button to get category labels of chart</span>
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
        let excelFileName = 'SampeB_4.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();

        const sb = [];
        //Create a workbook
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Get the chart
        let chart = sheet.Charts.get(0);

        //Get the cell range of the category labels
        let cr = chart.PrimaryCategoryAxis.CategoryLabels;
        for (let i = 0; i < cr.Count; i++) {
            sb.push(cr.Cells.get(i).Value);
        }
        
        // Dispose of the workbook object to free resources
        workbook.Dispose();
        // Create a Blob object
        const outputFile = 'GetCategoryLabels.txt';
        const modifiedFile = new Blob([sb.toString()], { type: "text/plain;charset=utf-8" });

        // Download the converted txt file
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
