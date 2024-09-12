<template>
  <span>Click the following button to set row color with conditional formatting in Excel file</span>
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
        let excelFileName='Template_Xls_4.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first worksheet
        const sheet = workbook.Worksheets.get(0);

        //Select the range that you want to format
        const dataRange = sheet.AllocatedRange;

        //Set conditional formatting
        const xcfs = sheet.ConditionalFormats.Add();
        xcfs.AddRange(dataRange);

        const format1 = xcfs.AddCondition();
        //Determines the cells to format
        format1.FirstFormula = "=MOD(ROW(),2)=0";
        //Set conditional formatting type
        format1.FormatType = wasmModule.ConditionalFormatType.Formula;
        //Set the color
        format1.BackColor = wasmModule.Color.get_LightSeaGreen();

        //Set the backcolor of the odd rows as Yellow
        const xcfs1 = sheet.ConditionalFormats.Add();
        xcfs1.AddRange(dataRange);

        const format2 = xcfs1.AddCondition();
        format2.FirstFormula = "=MOD(ROW(),2)=1";
        format2.FormatType = wasmModule.ConditionalFormatType.Formula;
        format2.BackColor = wasmModule.Color.get_Yellow();

        //Save result file
        const outputFileName = 'SetRowColorByConditionalFormat_out.xlsx';
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
