<template>
  <span>Click the following button to add trendline in a chart</span>
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
        let excelFileName = 'ChartSample2.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        
        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
       
        // Load the Excel file 
        workbook.LoadFromFile(excelFileName);

        //Get the first sheet
        let sheet = workbook.Worksheets.get(0);
        //select chart and set logarithmic trendline
        let chart = sheet.Charts.get(0);
        chart.ChartTitle = "Logarithmic Trendline";
        chart.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Logarithmic});

        //select chart and set moving_average trendline
        let chart1 = sheet.Charts.get(1);
        chart1.ChartTitle = "Moving Average Trendline";
        chart1.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Moving_Average});

        //select chart and set linear trendline
        let chart2 = sheet.Charts.get(2);
        chart2.ChartTitle = "Linear Trendline";
        chart2.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Linear});

        //select chart and set exponential trendline
        let chart3 = sheet.Charts.get(3);
        chart3.ChartTitle = "Exponential Trendline";
        chart3.Series.get(0).TrendLines.Add({type:spirexls.TrendLineType.Exponential});

        // Save the modified workbook 
        const outputFile = 'AddTrendline.xlsx';
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
