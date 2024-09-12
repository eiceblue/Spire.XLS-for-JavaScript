<template>
  <span>Click the following button to format cells with a style</span>
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
        let excelFileName='SampleB_2.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Create a style
        const style = workbook.Styles.Add("newStyle");
        //Set the shading color
        style.Color = wasmModule.Color.get_DarkGray();
        //Set the font color
        style.Font.Color = wasmModule.Color.get_White();
        //Set font name
        style.Font.FontName = "Times New Roman";
        //Set font size
        style.Font.Size = 12;
        //Set bold for the font
        style.Font.IsBold = true;
        //Set text rotation
        style.Rotation = 45;
        //Set alignment
        style.HorizontalAlignment = wasmModule.HorizontalAlignType.Center;
        style.VerticalAlignment = wasmModule.VerticalAlignType.Center;

        //Set the style for the specific range
        workbook.Worksheets.get(0).Range.get("A1:J1").CellStyleName = style.Name;

        //Save result file
        const outputFileName = 'FormatCellsWithStyle_out.xlsx';
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
