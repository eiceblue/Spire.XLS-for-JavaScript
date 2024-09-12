<template>
  <span>Click the following button to get the cropped position of a picture</span>
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
        let inputFileName='CroppedPositionOfPicture.xlsx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
         
        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({fileName: inputFileName});

        // Get the first worksheet
        let sheet1 = workbook.Worksheets.get(0);

        // Get the image from the first sheet
        let picture = sheet1.Pictures.get(0);

        // Get the cropped position
        let left = picture.Left;
        let top = picture.Top;
        let width = picture.Width;
        let height = picture.Height;

        // Create an array to save content
        let content = [];

        // Set string format for displaying
        let displayString = `Crop position: Left ${left}\r\nCrop position: Top ${top}\r\nCrop position: Width ${width}\r\nCrop position: Height ${height}`;

        // Add result string to content array
        content.push(displayString);

        const outputFileName = 'CroppedPositionOfPicture-out.txt';
        // Save them to a txt file
        writeTextToFile(content.join('\n'),outputFileName);

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
}
</script>