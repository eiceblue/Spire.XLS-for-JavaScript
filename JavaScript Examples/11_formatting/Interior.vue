<template>
  <span>Click the following button to set interior style of cell</span>
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

        //Create a workbook
        const workbook = wasmModule.Workbook.Create();

        //Initialize the workbook
        const sheet = workbook.Worksheets.get(0);

        //Specify the version
        workbook.Version = wasmModule.ExcelVersion.Version2007;

        //Define the number of the colors
        const maxColor = Object.keys(wasmModule.ExcelColors).length;

        //Create a random object
        for (let i = 2; i < 40; i++)
         {
            //Random backKnownColor
            const backKnownColor =wasmModule.ExcelColors.fromValue(Math.floor(Math.random() * (maxColor / 2)));

            //Add text
            sheet.Range.get("A1").Text = "Color Name";
            sheet.Range.get("B1").Text = "Red";
            sheet.Range.get("C1").Text = "Green";
            sheet.Range.get("D1").Text = "Blue";

            //Merge the sheet"E1-K1"
            sheet.Range.get("E1:K1").Merge();
            sheet.Range.get("E1:K1").Text = "Gradient";
            sheet.Range.get("A1:K1").Style.Font.IsBold = true;
            sheet.Range.get("A1:K1").Style.Font.Size = 11;

            //Set the text of color in sheetA-sheetD
            const colorName = backKnownColor;
            sheet.Range.get(`A${i}`).Text = colorName;
            sheet.Range.get(`B${i}`).NumberValue = workbook.GetPaletteColor(backKnownColor).R;
            sheet.Range.get(`C${i}`).NumberValue = workbook.GetPaletteColor(backKnownColor).G;
            sheet.Range.get(`D${i}`).NumberValue = workbook.GetPaletteColor(backKnownColor).B;

            //Merge the sheets
            sheet.Range.get(`E${i}:K${i}`).Merge();

            //Set the text of sheetE-sheetK
            sheet.Range.get(`E${i}:K${i}`).Text = colorName;

            //Set the interior of the color
            sheet.Range.get(`E${i}:K${i}`).Style.Interior.FillPattern = wasmModule.ExcelPatternType.Gradient;
            sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.BackKnownColor = backKnownColor;
            sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.ForeKnownColor = wasmModule.ExcelColors.White;
            sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.GradientStyle = wasmModule.GradientStyleType.Vertical;
            sheet.Range.get(`E${i}:K${i}`).Style.Interior.Gradient.GradientVariant = wasmModule.GradientVariantsType.ShadingVariants1;
        }

        //AutoFit Column
        sheet.AutoFitColumn(1);

        //Save result file
        const outputFileName = 'Interior_out.xlsx';
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
