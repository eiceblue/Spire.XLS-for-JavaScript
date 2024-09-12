<template>
  <span>Click the following button to copy shapes into other worksheet</span>
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
        // Create a new workbook object
        const workbook = wasmModule.Workbook.Create();
        // Load the Excel file 
        let sheet = workbook.Worksheets.get(0);

        // Create line shape
        let line = sheet.TypedLines.AddLine();
        line.Top = 50;
        line.Left = 30;
        line.Width = 30;
        line.Height = 50;
        line.BeginArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrowDiamond;
        line.EndArrowHeadStyle = wasmModule.ShapeArrowStyleType.LineArrow;

        let copySheet = workbook.Worksheets.get(1);
        // Copy the line into another sheet
        copySheet.TypedLines.AddCopy(line);

        // Create a button and then copy into another sheet
        let button = sheet.TypedRadioButtons.Add({ row: 5, column: 5, height: 20, width: 20 });
        copySheet.TypedRadioButtons.AddCopy(button);

        // Create a textbox and then copy into another sheet
        let textbox = sheet.TypedTextBoxes.AddTextBox(5, 7, 50, 100);
        copySheet.TypedTextBoxes.AddCopy(textbox);

        // Create a checkbox and then copy into another sheet
        let checkbox = sheet.TypedCheckBoxes.AddCheckBox(10, 1, 20, 20);
        copySheet.TypedCheckBoxes.AddCopy(checkbox);

        // Create a combobox and then copy into another sheet
        sheet.Range.get("A14").Value = "1";
        sheet.Range.get("A15").Value = "2";
        let comboBoxes = sheet.TypedComboBoxes.AddComboBox(10, 5, 30, 30);
        comboBoxes.ListFillRange = sheet.Range.get("A14:A15");
        copySheet.TypedComboBoxes.AddCopy(comboBoxes);

        // Save the modified workbook 
        const outputFile = 'CopyShapes.xlsx';
        workbook.SaveToFile(outputFile);
        // Dispose of the workbook object to free resources
        workbook.Dispose();

        // Read the saved Excel file from the virtual file system and convert it to a Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFile);
        const modifiedFile = new Blob([modifiedFileArray], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

        // Download the converted Excel file
        downloadName.value = outputFile;
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
