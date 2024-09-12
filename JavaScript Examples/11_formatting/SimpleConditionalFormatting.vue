<template>
  <span>Click the following button to add conditional formatting in an existing excel workbook</span>
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
        let excelFileName='ConditionalFormatting.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        function AddConditionalFormattingForExistingSheet(sheet) {
            sheet.AllocatedRange.RowHeight = 15;
            sheet.AllocatedRange.ColumnWidth = 16;

            // Create conditional formatting rule
            const xcfs1 = sheet.ConditionalFormats.Add();
            xcfs1.AddRange(sheet.Range.get("A1:D1"));
            const cf1 = xcfs1.AddCondition();
            cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
            cf1.FirstFormula = "150";
            cf1.Operator = wasmModule.ComparisonOperatorType.Greater;
            cf1.FontColor = wasmModule.Color.get_Red();
            cf1.BackColor = wasmModule.Color.get_LightBlue();

            const xcfs2 = sheet.ConditionalFormats.Add();
            xcfs2.AddRange(sheet.Range.get("A2:D2"));
            const cf2 = xcfs2.AddCondition();
            cf2.FormatType = wasmModule.ConditionalFormatType.CellValue;
            cf2.FirstFormula = "300";
            cf2.Operator = wasmModule.ComparisonOperatorType.Less;
            // Set border color
            cf2.LeftBorderColor = wasmModule.Color.get_Pink();
            cf2.RightBorderColor = wasmModule.Color.get_Pink();
            cf2.TopBorderColor = wasmModule.Color.get_DeepSkyBlue();
            cf2.BottomBorderColor = wasmModule.Color.get_DeepSkyBlue();
            cf2.LeftBorderStyle = wasmModule.LineStyleType.Medium;
            cf2.RightBorderStyle = wasmModule.LineStyleType.Thick;
            cf2.TopBorderStyle = wasmModule.LineStyleType.Double;
            cf2.BottomBorderStyle = wasmModule.LineStyleType.Double;

            // Add data bars
            const xcfs3 = sheet.ConditionalFormats.Add();
            xcfs3.AddRange(sheet.Range.get("A3:D3"));
            const cf3 = xcfs3.AddCondition();
            cf3.FormatType = wasmModule.ConditionalFormatType.DataBar;
            cf3.DataBar.BarColor = wasmModule.Color.get_CadetBlue();

            // Add icon sets
            const xcfs4 = sheet.ConditionalFormats.Add();
            xcfs4.AddRange(sheet.Range.get("A4:D4"));
            const cf4 = xcfs4.AddCondition();
            cf4.FormatType = wasmModule.ConditionalFormatType.IconSet;
            cf4.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights1;

            // Add color scales
            const xcfs5 = sheet.ConditionalFormats.Add();
            xcfs5.AddRange(sheet.Range.get("A5:D5"));
            const cf5 = xcfs5.AddCondition();
            cf5.FormatType = wasmModule.ConditionalFormatType.ColorScale;

            // Highlight duplicate values in range "A6:D6" with BurlyWood color
            const xcfs6 = sheet.ConditionalFormats.Add();
            xcfs6.AddRange(sheet.Range.get("A6:D6"));
            const cf6 = xcfs6.AddCondition();
            cf6.FormatType = wasmModule.ConditionalFormatType.DuplicateValues;
            cf6.BackColor = wasmModule.Color.get_BurlyWood();
        }
    
        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        // Get the first sheet
        const oldSheet = workbook.Worksheets.get(0);
        AddConditionalFormattingForExistingSheet(oldSheet);

        //Save result file
        const outputFileName = 'SimpleConditionalFormatting_out.xlsx';
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
