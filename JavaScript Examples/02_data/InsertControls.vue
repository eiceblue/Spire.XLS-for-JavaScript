<template>
  <span>Click the following button to insert controls in Excel file</span>
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

        // Load the sample file into the virtual file system (VFS)
        let excelFileName = "InsertControls.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();

        // Load an existing Excel from the virtual file system
        workbook.LoadFromFile(excelFileName);

        // Get the first worksheet
        let ws = workbook.Worksheets.get(0);

        //Add a textbox
        let textbox = ws.TextBoxes.AddTextBox(9, 2, 25, 100);
        textbox.Text = "Hello World";

        //Add a checkbox
        let cb = ws.CheckBoxes.AddCheckBox(11, 2, 15, 100);
        cb.CheckState = wasmModule.CheckState.Checked;
        cb.Text = "Check Box 1";

        //Add a RadioButton
        let rb = ws.RadioButtons.Add({
          row: 13,
          column: 2,
          height: 15,
          width: 100,
        });
        rb.Text = "Option 1";

        // Add a combox
        let cbx = ws.ComboBoxes.AddComboBox(15, 2, 15, 100);
        cbx.ListFillRange = ws.Range.get("A36:A42");

        // Define the output file name
        const outputFileName = "InsertControls_out.xlsx";

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
