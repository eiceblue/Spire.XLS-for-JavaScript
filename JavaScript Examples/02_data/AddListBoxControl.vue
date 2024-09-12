<template>
  <span>Click the following button to add listbox control in Excel file</span>
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
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Add an empty sheet to the workbook
        workbook.CreateEmptySheets(1);

        // Get the first worksheet
        let sheet = workbook.Worksheets.get(0);

        // Set text for cells
        sheet.Range.get("A7").Text = "Beijing";
        sheet.Range.get("A8").Text = "New York";
        sheet.Range.get("A9").Text = "ChengDu";
        sheet.Range.get("A10").Text = "Paris";
        sheet.Range.get("A11").Text = "Boston";
        sheet.Range.get("A12").Text = "London";

        sheet.Range.get("C13").Text = "City :";
        sheet.Range.get("C13").Style.Font.IsBold = true;

        // Add listbox control
        let listBox = sheet.ListBoxes.AddListBox(13, 4, 100, 80);
        listBox.SelectionType = wasmModule.SelectionType.Single;
        listBox.SelectedIndex = 2;
        listBox.Display3DShading = true;
        listBox.ListFillRange = sheet.Range.get("A7:A12");

        // Define the output file name
        const outputFileName = "AddListBoxControl.xlsx";

        // Save the workbook to the specified path
        workbook.SaveToFile({
          fileName: outputFileName,
          version: wasmModule.ExcelVersion.Version2010,
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
