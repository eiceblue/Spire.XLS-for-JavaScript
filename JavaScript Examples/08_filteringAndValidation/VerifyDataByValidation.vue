<template>
  <span>Click the following button to verify data by the validation in worksheet</span>
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
      if (wasmModule) {

        let inputFileName='VerifyDataByValidation.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        //Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Cell B4 has the Decimal Validation
        let cell = sheet.Range.get("B4");

        //Get the valditation of this cell
        let validation = cell.DataValidation;

        //Get the specified data range
        let minimum = parseFloat(validation.Formula1);
        let maximum = parseFloat(validation.Formula2);

        //Create StringBuilder to save
        let content = [];

        //Set different numbers for the cell
        for(let i=5; i<100; i+=40) {
            cell.NumberValue = i;
            let result = null;
            //Verify
            if(cell.NumberValue < minimum || cell.NumberValue > maximum) {
                //Set string format for displaying
                result = `Is input ${i} a valid value for this Cell: false`;
            } else {
                //Set string format for displaying
                result = `Is input ${i} a valid value for this Cell: true`;
            }
            //Add result string to StringBuilder
            content.push(result);
        }

        const outputFileName = 'VerifyDataByValidation-out.txt';
        
        // Save the content to the specified path
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));

        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
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
</script>
