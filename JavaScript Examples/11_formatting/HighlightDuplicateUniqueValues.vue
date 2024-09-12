<template>
  <span>Click the following button to highlight duplicate and unique values in Excel file</span>
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

        //Use conditional formatting to highlight duplicate values in range "C2:C10" with IndianRed color.
        const xcfs = sheet.ConditionalFormats.Add();
        xcfs.AddRange(sheet.Range.get("C2:C10"));
        const format1 = xcfs.AddCondition();
        format1.FormatType = wasmModule.ConditionalFormatType.DuplicateValues;
        format1.BackColor = wasmModule.Color.get_IndianRed();

        //Use conditional formatting to highlight unique values in range "C2:C10" with Yellow color.
        const xcfs1 = sheet.ConditionalFormats.Add();
        xcfs1.AddRange(sheet.Range.get("C2:C10"));
        const format2 = xcfs1.AddCondition();
        format2.FormatType = wasmModule.ConditionalFormatType.UniqueValues;
        format2.BackColor = wasmModule.Color.get_Yellow();

        //Save result file
        const outputFileName = 'HighlightDuplicateUniqueValues_out.xlsx';
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
