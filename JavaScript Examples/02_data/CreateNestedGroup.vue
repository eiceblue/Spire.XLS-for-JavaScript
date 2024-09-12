<template>
  <span
    >Click the following button to create nested group in Excel file</span
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
        let sheet = workbook.Worksheets.get(0);

        //Set the style.
        let style = workbook.Styles.Add("style");
        style.Font.Color = wasmModule.Color.get_CadetBlue();
        style.Font.IsBold = true;

        //Set the summary rows appear above detail rows
        sheet.PageSetup.IsSummaryRowBelow = false;

        //Insert sample data to cells
        sheet.Range.get("A1").Value = "Project plan for project X";
        sheet.Range.get("A1").CellStyleName = style.Name;

        sheet.Range.get("A3").Value = "Set up";
        sheet.Range.get("A3").CellStyleName = style.Name;
        sheet.Range.get("A4").Value = "Task 1";
        sheet.Range.get("A5").Value = "Task 2";
        sheet.Range.get("A4:A5").BorderAround({
          borderLine: wasmModule.LineStyleType.Thin,
        });
        sheet.Range.get("A4:A5").BorderInside({
          borderLine: wasmModule.LineStyleType.Thin,
        });

        sheet.Range.get("A7").Value = "Launch";
        sheet.Range.get("A7").CellStyleName = style.Name;
        sheet.Range.get("A8").Value = "Task 1";
        sheet.Range.get("A9").Value = "Task 2";
        sheet.Range.get("A8:A9").BorderAround({
          borderLine: wasmModule.LineStyleType.Thin,
        });
        sheet.Range.get("A8:A9").BorderInside({
          borderLine: wasmModule.LineStyleType.Thin,
        });

        //Group the rows that you want to group.
        sheet.GroupByRows(2, 9, false);
        sheet.GroupByRows(4, 5, false);
        sheet.GroupByRows(8, 9, false);

        // Define the output file name
        const outputFileName = "CreateNestedGroup.xlsx";

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
