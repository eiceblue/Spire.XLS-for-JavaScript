<template>
  <span
    >Click the following button to copy data with style in worksheet</span
  >
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      const wasmModule = window.wasmModule; 
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Get the first worksheet
        let worksheet = workbook.Worksheets.get(0);

        // Set the values for some cells
        let cells = worksheet.Range.get("A1:J50");
        for (let i = 1; i <= 10; i++) {
          for (let j = 1; j <= 8; j++) {
            let text = i - 1 + "," + (j - 1);
            cells.get({ row: i, column: j }).Text = text;
          }
        }
        // Get a source range (A1:D3)
        let srcRange = worksheet.Range.get("A1:D3");

        // Create a style object
        let style = workbook.Styles.Add("style");

        // Specify the font attribute
        style.Font.FontName = "Calibri";

        // Specify the shading color
        style.Font.Color = wasmModule.Color.get_Red();

        // Specify the border attributes
        style.Borders.get(wasmModule.BordersLineType.EdgeTop).LineStyle =
          wasmModule.LineStyleType.Thin;
        style.Borders.get(wasmModule.BordersLineType.EdgeTop).Color =
          wasmModule.Color.get_Blue();
        style.Borders.get(wasmModule.BordersLineType.EdgeBottom).LineStyle =
          wasmModule.LineStyleType.Thin;
        style.Borders.get(wasmModule.BordersLineType.EdgeBottom).Color =
          wasmModule.Color.get_Blue();
        style.Borders.get(wasmModule.BordersLineType.EdgeLeft).LineStyle =
          wasmModule.LineStyleType.Thin;
        style.Borders.get(wasmModule.BordersLineType.EdgeLeft).Color =
          wasmModule.Color.get_Blue();
        style.Borders.get(wasmModule.BordersLineType.EdgeRight).LineStyle =
          wasmModule.LineStyleType.Thin;
        style.Borders.get(wasmModule.BordersLineType.EdgeRight).Color =
          wasmModule.Color.get_Blue();
        srcRange.CellStyleName = style.Name;

        // Set the destination range
        let destRange = worksheet.Range.get("A12:D14");

        // Copy the range data with style
        srcRange.Copy(destRange, true, true);

        // Define the output file name
        const outputFileName = "CopyDataWithStyle.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
        
        // Clean up resources
        workbook.Dispose();
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
