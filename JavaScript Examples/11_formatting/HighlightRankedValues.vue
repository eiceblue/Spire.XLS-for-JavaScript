<template>
  <span>Click the following button to highlight top and bottom ranked values in Excel file</span>
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

        //Apply conditional formatting to range "D2:D10" to highlight the top 2 values.
        const xcfs = sheet.ConditionalFormats.Add();
        xcfs.AddRange(sheet.Range.get("D2:D10"));
        const format1 = xcfs.AddTopBottomCondition(wasmModule.TopBottomType.Top, 2);
        format1.FormatType = wasmModule.ConditionalFormatType.TopBottom;
        format1.BackColor = wasmModule.Color.get_Red();

        //Apply conditional formatting to range "E2:E10" to highlight the bottom 2 values.
        const xcfs1 = sheet.ConditionalFormats.Add();
        xcfs1.AddRange(sheet.Range.get("E2:E10"));
        const format2 = xcfs1.AddTopBottomCondition(wasmModule.TopBottomType.Bottom, 2);
        format2.FormatType = wasmModule.ConditionalFormatType.TopBottom;
        format2.BackColor = wasmModule.Color.get_ForestGreen();

        //Save result file
        const outputFileName = 'HighlightRankedValues_out.xlsx';
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
