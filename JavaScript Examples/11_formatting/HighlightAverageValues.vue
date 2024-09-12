<template>
  <span>Click the following button to highlight below and above average values in Excel file</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref} from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Input file
        let excelFileName='ConditionallyFormatDate.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first worksheet.
        const sheet = workbook.Worksheets.get(0);

        //Add conditional format.
        const format1 = sheet.ConditionalFormats.Add();
        //Set the cell range to apply the formatting.
        format1.AddRange(sheet.Range.get("E2:E10"));
        //Add below average condition.
        const cf1 = format1.AddAverageCondition(wasmModule.AverageType.Below);
        //Highlight cells below average values.
        cf1.BackColor = wasmModule.Color.get_SkyBlue();

        //Add conditional format.
        const format2 = sheet.ConditionalFormats.Add();
        //Set the cell range to apply the formatting.
        format2.AddRange(sheet.Range.get("E2:E10"));
        //Add above average condition.
        const cf2 = format2.AddAverageCondition(wasmModule.AverageType.Above);
        //Highlight cells above average values.
        cf2.BackColor = wasmModule.Color.get_Orange();

        //Save result file
        const outputFileName = 'HighlightAverageValues_out.xlsx';
        workbook.SaveToFile({fileName: outputFileName, version:wasmModule.ExcelVersion.Version2010});

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        // Download the result file
        downloadName.value = outputFileName;
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
