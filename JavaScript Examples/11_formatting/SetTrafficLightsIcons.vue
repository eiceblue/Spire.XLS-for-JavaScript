<template>
  <span>Click the following button to set traffic lights icons in Excel file</span>
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

        // Create a workbook
        const workbook = wasmModule.Workbook.Create();

        // Add a worksheet
        const sheet = workbook.Worksheets.get(0);

        // Add some data to the Excel sheet cell range and set the format for them
        sheet.Range.get("A1").Text = "Traffic Lights";
        sheet.Range.get("A2").NumberValue = 0.95;
        sheet.Range.get("A2").NumberFormat = "0%";
        sheet.Range.get("A3").NumberValue = 0.5;
        sheet.Range.get("A3").NumberFormat = "0%";
        sheet.Range.get("A4").NumberValue = 0.1;
        sheet.Range.get("A4").NumberFormat = "0%";
        sheet.Range.get("A5").NumberValue = 0.9;
        sheet.Range.get("A5").NumberFormat = "0%";
        sheet.Range.get("A6").NumberValue = 0.7;
        sheet.Range.get("A6").NumberFormat = "0%";
        sheet.Range.get("A7").NumberValue = 0.6;
        sheet.Range.get("A7").NumberFormat = "0%";

        // Set the height of row and width of column for Excel cell range
        sheet.AllocatedRange.RowHeight = 20;
        sheet.AllocatedRange.ColumnWidth = 25;

        // Add a conditional formatting
        const conditional = sheet.ConditionalFormats.Add();
        conditional.AddRange(sheet.AllocatedRange);
        const format1 = conditional.AddCondition();

        // Add a conditional formatting of cell range and set its type to CellValue
        format1.FormatType = wasmModule.ConditionalFormatType.CellValue;
        format1.FirstFormula = "300";
        format1.Operator = wasmModule.ComparisonOperatorType.Less;
        format1.FontColor = wasmModule.Color.get_Black();
        format1.BackColor = wasmModule.Color.get_LightSkyBlue();

        // Add a conditional formatting of cell range and set its type to IconSet
        conditional.AddRange(sheet.AllocatedRange);
        const format = conditional.AddCondition();
        format.FormatType = wasmModule.ConditionalFormatType.IconSet;
        format.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights1;

        //Save result file
        const outputFileName = 'SetTrafficLightsIcons_out.xlsx';
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
