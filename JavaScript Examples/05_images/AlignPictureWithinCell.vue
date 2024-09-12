<template>
  <span>Click the following button to align the picture within cell</span>
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

        let inputFileName='SpireXls.png';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Get the first worksheet in the workbook
        let sheet = workbook.Worksheets.get(0);
        // Set the text in cell A1
        sheet.Range.get("A1").Text = "Align Picture Within A Cell:";
        // Set the vertical alignment of cell A1 to top
        sheet.Range.get("A1").Style.VerticalAlignment = wasmModule.VerticalAlignType.Top;
        // Insert an image at the specific cell (1, 1)
        let picture = sheet.Pictures.Add({topRow:1, leftColumn:1, fileName:inputFileName});
        // Adjust the column width and row height so that the cell can contain the picture
        sheet.Columns.get(0).ColumnWidth = 40;
        sheet.Rows.get(0).RowHeight =200;
        // Set the horizontal offset of the image within the cell to 100
        picture.LeftColumnOffset = 100;
        // Set the vertical offset of the image within the cell to 25
        picture.TopRowOffset = 25;
        const outputFileName = 'AlignPictureWithinCell.xlsx';
        // Save the workbook to the specified file
        workbook.SaveToFile({fileName:outputFileName,version:wasmModule.ExcelVersion.Version2010});
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

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