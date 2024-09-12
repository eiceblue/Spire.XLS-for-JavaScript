<template>
  <span>Click the following button to format text within a cell</span>
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
        let excelFileName='FontStyles.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);
      
        //Set font style
        sheet.Range.get("B1").Style.Font.FontName = "Comic Sans MS";
        sheet.Range.get("B2:D2").Style.Font.FontName = "Corbel";
        sheet.Range.get("B3:D7").Style.Font.FontName = "Aleo";
      
        //Set font size
        sheet.Range.get("B1").Style.Font.Size = 45;
        sheet.Range.get("B2:D3").Style.Font.Size = 25;
        sheet.Range.get("B3:D7").Style.Font.Size = 12;
      
        //Set excel cell data to be bold
        sheet.Range.get("B2:D2").Style.Font.IsBold = true;
      
        //Set excel cell data to be underline
        sheet.Range.get("B3:B7").Style.Font.Underline = wasmModule.FontUnderlineType.Single;
      
        //Set excel cell data color
        sheet.Range.get("B1").Style.Font.Color = wasmModule.Color.get_CornflowerBlue();
        sheet.Range.get("B2:D2").Style.Font.Color = wasmModule.Color.get_CadetBlue();
        sheet.Range.get("B3:D7").Style.Font.Color = wasmModule.Color.get_Firebrick();
      
        //Set excel cell data to be italic
        sheet.Range.get("B3:D7").Style.Font.IsItalic = true;
      
        //Add strikethrough
        sheet.Range.get("D3").Style.Font.IsStrikethrough = true;
        sheet.Range.get("D7").Style.Font.IsStrikethrough = true;

        //Save result file
        const outputFileName = 'FontStyles_out.xlsx';
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
