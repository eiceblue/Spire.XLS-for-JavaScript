<template>
  <span>Click the following button to apply color scales to data range in Excel file</span>
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

        //Create a workbook.
        const workbook = wasmModule.Workbook.Create();

        //Get the first worksheet.
        const sheet = workbook.Worksheets.get(0);

        //Insert data to cell range from A1 to C4.
        sheet.Range.get("A1").NumberValue = 582;
        sheet.Range.get("A2").NumberValue = 234;
        sheet.Range.get("A3").NumberValue = 314;
        sheet.Range.get("A4").NumberValue = 50;
        sheet.Range.get("B1").NumberValue = 150;
        sheet.Range.get("B2").NumberValue = 894;
        sheet.Range.get("B3").NumberValue = 560;
        sheet.Range.get("B4").NumberValue = 900;
        sheet.Range.get("C1").NumberValue = 134;
        sheet.Range.get("C2").NumberValue = 700;
        sheet.Range.get("C3").NumberValue = 920;
        sheet.Range.get("C4").NumberValue = 450;
        sheet.AllocatedRange.RowHeight = 15;
        sheet.AllocatedRange.ColumnWidth = 17;

        //Add color scales.
        const xcfs = sheet.ConditionalFormats.Add();
        xcfs.AddRange(sheet.AllocatedRange);
        const format = xcfs.AddCondition();
        format.FormatType = wasmModule.ConditionalFormatType.ColorScale;

        //Save result file
        const outputFileName = 'ApplyColorScalesToDataRange_out.xlsx';
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
