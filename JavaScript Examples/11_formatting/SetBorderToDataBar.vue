<template>
  <span>Click the following button to add border to databar</span>
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
        let excelFileName='Template_Xls_9.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);

        //Get the databar format 
        const xcfs = sheet.ConditionalFormats.get(0);
        const cf = xcfs.get(0);
        const dataBar1 = cf.DataBar;
        dataBar1.BarBorder.Type = wasmModule.DataBarBorderType.DataBarBorderSolid;
        dataBar1.BarBorder.Color = wasmModule.Color.get_Red();

        //Set to new data bar
        sheet.Range.get("E1").NumberValue = 200;
        const xcfs2 = sheet.ConditionalFormats.Add();
        xcfs2.AddRange(sheet.Range.get("E1"));
        const cf2 = xcfs2.AddCondition();
        cf2.FormatType = wasmModule.ConditionalFormatType.DataBar;
        cf2.DataBar.BarBorder.Type = wasmModule.DataBarBorderType.DataBarBorderSolid;
        cf2.DataBar.BarBorder.Color = wasmModule.Color.get_Red();
        cf2.DataBar.BarColor = wasmModule.Color.get_GreenYellow();

        //Save result file
        const outputFileName = 'SetBorderToDataBar_out.xlsx';
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
