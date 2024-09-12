<template>
  <span>Click the following button to get color Argb data</span>
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
        let excelFileName='templateAz.xlsx';
        await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Load the document
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile({fileName: excelFileName});

        //Get the first sheet
        const sheet = workbook.Worksheets.get(0);

        const strB = [];

        //Get font color
        const color1 = sheet.Range.get("B2").Style.Font.Color;

        //Read ARGB data of Color
        strB.push(`The font color of B2: ARGB=(${color1.A},${color1.R},${color1.G},${color1.B})`);

        const color2 = sheet.Range.get("B3").Style.Font.Color;
        strB.push(`The font color of B3: ARGB=(${color2.A},${color2.R},${color2.G},${color2.B})`);

        const color3 = sheet.Range.get("B4").Style.Font.Color;
        strB.push(`The font color of B4: ARGB=(${color3.A},${color3.R},${color3.G},${color3.B})`);

        //Save result file
        const outputFileName = 'GetColorArgbData_out.txt';
        FS.writeFile(outputFileName, strB.join('\n'))

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'text/plain'});

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
