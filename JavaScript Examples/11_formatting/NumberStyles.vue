<template>
  <span>Click the following button to set number format</span>
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
        let excelFileName='NumberStyles.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Initialize the workbook
        const sheet = workbook.Worksheets.get(0);

        //Input a number value for the specified cell and set the number format
        sheet.Range.get("B10").Text = "NUMBER FORMATTING";
        sheet.Range.get("B10").Style.Font.IsBold = true;

        sheet.Range.get("B13").Text = "0";
        sheet.Range.get("C13").NumberValue = 1234.5678;
        sheet.Range.get("C13").NumberFormat = "0";

        sheet.Range.get("B14").Text = "0.00";
        sheet.Range.get("C14").NumberValue = 1234.5678;
        sheet.Range.get("C14").NumberFormat = "0.00";

        sheet.Range.get("B15").Text = "#,##0.00";
        sheet.Range.get("C15").NumberValue = 1234.5678;
        sheet.Range.get("C15").NumberFormat = "#,##0.00";

        sheet.Range.get("B16").Text = "$#,##0.00";
        sheet.Range.get("C16").NumberValue = 1234.5678;
        sheet.Range.get("C16").NumberFormat = "$#,##0.00";

        sheet.Range.get("B17").Text = "0;[Red]-0";
        sheet.Range.get("C17").NumberValue = -1234.5678;
        sheet.Range.get("C17").NumberFormat = "0;[Red]-0";

        sheet.Range.get("B18").Text = "0.00;[Red]-0.00";
        sheet.Range.get("C18").NumberValue = -1234.5678;
        sheet.Range.get("C18").NumberFormat = "0.00;[Red]-0.00";

        sheet.Range.get("B19").Text = "#,##0;[Red]-#,##0";
        sheet.Range.get("C19").NumberValue = -1234.5678;
        sheet.Range.get("C19").NumberFormat = "#,##0;[Red]-#,##0";

        sheet.Range.get("B20").Text = "#,##0.00;[Red]-#,##0.00";
        sheet.Range.get("C20").NumberValue = -1234.5678;
        sheet.Range.get("C20").NumberFormat = "#,##0.00;[Red]-#,##0.00";

        sheet.Range.get("B21").Text = "0.00E+00";
        sheet.Range.get("C21").NumberValue = 1234.5678;
        sheet.Range.get("C21").NumberFormat = "0.00E+00";

        sheet.Range.get("B22").Text = "0.00%";
        sheet.Range.get("C22").NumberValue = 1234.5678;
        sheet.Range.get("C22").NumberFormat = "0.00%";

        sheet.Range.get("B13:B22").Style.KnownColor = wasmModule.ExcelColors.Gray25Percent;

        //AutoFit Column
        sheet.AutoFitColumn(2);
        sheet.AutoFitColumn(3);

        //Save result file
        const outputFileName = 'NumberStyles_out.xlsx';
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
