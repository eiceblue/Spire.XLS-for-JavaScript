<template>
  <span>Click the following button to write formulas</span>
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

        //Create a workbook.
		const workbook = wasmModule.Workbook.Create();
		
		//Get the first sheet
		const sheet = workbook.Worksheets.get(0);

		let currentRow = 1;
		let currentFormula = "";

		//Set column width 
		sheet.SetColumnWidth(1, 32);
		sheet.SetColumnWidth(2, 16);
		sheet.SetColumnWidth(3, 16);		
		
        //Set values
		sheet.Range.get({row:currentRow, column:1}).Value = "Examples of formulas :";
		currentRow += 2;

		sheet.Range.get({row:currentRow, column:1}).Value = "Test data:";

		let range = sheet.Range.get("A1");
		range.Style.Font.IsBold = true;
		range.Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		range.Style.KnownColor = wasmModule.ExcelColors.LightGreen1;
		range.Style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle = wasmModule.LineStyleType.Medium;

		sheet.Range.get({row:currentRow, column:2}).NumberValue = 7.3;
		sheet.Range.get({row:currentRow, column:3}).NumberValue = 5;
		sheet.Range.get({row:currentRow, column:4}).NumberValue = 8.2;
		sheet.Range.get({row:currentRow, column:5}).NumberValue = 4;
		sheet.Range.get({row:currentRow, column:6}).NumberValue = 3;
		sheet.Range.get({row:currentRow, column:7}).NumberValue = 11.3;

		currentRow += 1;

		sheet.Range.get({row:currentRow, column:1}).Value = "Formulas";
		sheet.Range.get({row:currentRow, column:2}).Value = "Results";
		range = sheet.Range.get({row:currentRow, column:1, lastRow:currentRow, lastColumn:2});
		range.Style.Font.IsBold = true;
		range.Style.KnownColor = wasmModule.ExcelColors.LightGreen1;
		range.Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		range.Style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle = wasmModule.LineStyleType.Medium;

		currentFormula = "=\"hello\"";
		currentRow += 1;

		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = "=\"hello\"";
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		sheet.Range.get({row:currentRow, column:3}).Formula = "=\"" + '\u4f60\u597d' + "\"";

		currentFormula = "=300";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=3389.639421";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=false";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=1+2+3+4+5-6-7+8-9";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=33*3/4-2+10";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=Sheet1!$B$3";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=AVERAGE(Sheet1!$D$3:G$3)";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=Count(3,5,8,10,2,34)";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;

		currentFormula = "=NOW()";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Style.NumberFormat = "yyyy-MM-DD";

		currentFormula = "=SECOND(11)";
		currentRow += 1;
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=MINUTE(12)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=MONTH(9)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=DAY(10)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=TIME(4,5,7)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=DATE(6,4,2)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=RAND()";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=HOUR(12)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=MOD(5,3)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=WEEKDAY(3)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=YEAR(23)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=NOT(true)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=OR(true)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=AND(TRUE)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=VALUE(30)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=LEN(\"world\")";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=MID(\"world\",4,2)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=ROUND(7,3)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=SIGN(4)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=INT(200)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=ABS(-1.21)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=LN(15)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=EXP(20)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=SQRT(40)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=PI()";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=COS(9)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=SIN(45)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=MAX(10,30)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=MIN(5,7)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=AVERAGE(12,45)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=SUM(18,29)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=IF(4,2,2)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

		currentFormula = "=SUBTOTAL(3,Sheet1!B2:E3)";
		sheet.Range.get({row:currentRow, column:1}).NumberFormat="@";
		sheet.Range.get({row:currentRow, column:1}).Text = currentFormula;
		sheet.Range.get({row:currentRow, column:2}).Formula = currentFormula;
		currentRow += 1;

        //Save result file
        const outputFileName = 'WriteFormulas_out.xlsx';
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
