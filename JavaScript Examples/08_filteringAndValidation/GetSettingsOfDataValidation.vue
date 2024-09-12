<template>
  <span>Click the following button to get settings of data validation in worksheet</span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {

        let inputFileName='GetSettingsOfDataValidation.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        //Get the first worksheet
        let worksheet = workbook.Worksheets.get(0);
        
        //Cell B4 has the Decimal Validation
        let cell = worksheet.Range.get("B4");

        //Get the validation of this cell
        let validation = cell.DataValidation;

        //Get the settings
        let allowType = validation.AllowType.toString();
        let data = validation.CompareOperator.toString();
        let minimum = validation.Formula1.toString();
        let maximum = validation.Formula2.toString();
        let ignoreBlank = validation.IgnoreBlank.toString();

        //Create an array to save the content
        let content = [];

        //Set string format for displaying
        let result = `Settings of Validation: \r\nAllow Type: ${allowType}\r\nData: ${data}\r\nMinimum: ${minimum}\r\nMaximum: ${maximum}\r\nIgnoreBlank: ${ignoreBlank}`;

        //Add result string to the content array
        content.push(result);

        const outputFileName = 'GetSettingsOfDataValidation-out.txt';

        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));
        
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'text/plain'});

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
function writeTextToFile(text, filename) {
    FS.writeFile(filename, text, (err) => {
        if (err) throw err;
        console.log('The file has been saved!');
    });
};
</script>
