<template>
  <span>Click the following button to add runtime conditional formatting</span>
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
        let excelFileName='ConditionalFormatRuntime.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

		function  AddComparisonRule1(sheet) {
		//Create conditional formatting rule
		const xcfs1 = sheet.ConditionalFormats.Add();
		xcfs1.AddRange(sheet.Range.get("A1:D1"));
		const cf1 = xcfs1.AddCondition();
		cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
		cf1.FirstFormula = "150";
		cf1.Operator = wasmModule.ComparisonOperatorType.Greater;
		cf1.FontColor = wasmModule.Color.get_Red();
		cf1.BackColor = wasmModule.Color.get_LightBlue();
		}
	
		function  AddComparisonRule2(sheet) {
			const xcfs2 = sheet.ConditionalFormats.Add();
			xcfs2.AddRange(sheet.Range.get("A2:D2"));
			const cf2 = xcfs2.AddCondition();
			cf2.FormatType = wasmModule.ConditionalFormatType.CellValue;
			cf2.FirstFormula = "500";
			cf2.Operator = wasmModule.ComparisonOperatorType.Less;
			//Set border color
			cf2.LeftBorderColor = wasmModule.Color.get_Pink();
			cf2.RightBorderColor = wasmModule.Color.get_Pink();
			cf2.TopBorderColor = wasmModule.Color.get_DeepSkyBlue();
			cf2.BottomBorderColor = wasmModule.Color.get_DeepSkyBlue();
			cf2.LeftBorderStyle = wasmModule.LineStyleType.Medium;
			cf2.RightBorderStyle = wasmModule.LineStyleType.Thick;
			cf2.TopBorderStyle = wasmModule.LineStyleType.Double;
			cf2.BottomBorderStyle = wasmModule.LineStyleType.Double;
		}
	
		function  AddComparisonRule3(sheet) {
			//Create conditional formatting rule
			const xcfs1 = sheet.ConditionalFormats.Add();
			xcfs1.AddRange(sheet.Range.get("A3:D3"));
			const cf1 = xcfs1.AddCondition();
			cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
			cf1.FirstFormula = "300";
			cf1.SecondFormula = "500";
			cf1.Operator = wasmModule.ComparisonOperatorType.Between;
			cf1.BackColor = wasmModule.Color.get_Yellow();
		}
	
		function  AddComparisonRule4(sheet) {
			//Create conditional formatting rule
			const xcfs1 = sheet.ConditionalFormats.Add();
			xcfs1.AddRange(sheet.Range.get("A4:D4"));
			const cf1 = xcfs1.AddCondition();
			cf1.FormatType = wasmModule.ConditionalFormatType.CellValue;
			cf1.FirstFormula = "100";
			cf1.SecondFormula = "200";
			cf1.Operator = wasmModule.ComparisonOperatorType.NotBetween;
			//Set fill pattern type
			cf1.FillPattern = wasmModule.ExcelPatternType.ReverseDiagonalStripe;
			//Set foreground color
			cf1.Color = wasmModule.Color.FromArgb(255, 255, 0);
			//Set background color
			cf1.BackColor = wasmModule.Color.FromArgb(0, 255, 255);
		}
				
        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
		const sheet = workbook.Worksheets.get(0);
		
	
		AddComparisonRule1(sheet);
		AddComparisonRule2(sheet);
		AddComparisonRule3(sheet);
		AddComparisonRule4(sheet);

        //Save result file
        const outputFileName = 'ConditionalFormatRuntime_out.xlsx';
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
