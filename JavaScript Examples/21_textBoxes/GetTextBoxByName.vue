<template>
  <span>Click the following button to get TextBox by name in worksheet</span>
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
        // Load the fonts
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        //Get the default first worksheet
        let sheet = workbook.Worksheets.get(0);

        //Insert a TextBox
        sheet.Range.get("A2").Text = "Name：";
        let textBox = sheet.TextBoxes.AddTextBox(2, 2, 18, 65);

        //Set the name
        textBox.Name = "FirstTextBox";

        //Set string text for TextBox
        textBox.Text =
          "Spire.XLS for .NET is a professional Excel .NET component that can be used to any type of .NET 2.0, 3.5, 4.0 or 4.5 framework application, both ASP.NET web sites and Windows Forms application.";

        //Get the TextBox by the name
        let FindTextBox = sheet.TextBoxes.get("FirstTextBox");

        //Get the TextBox text
        let text = FindTextBox.Text;

        //Create content array to save
        let content = [];

        //Set string format for displaying
        let result = `The text of "${textBox.Name}" is: ${text}`;

        //Add result string to content array
        content.push(result);

        let outputFileName = "GetTextBoxByName_output.xlsx";
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));

        // Dispose of the workbook object to release resources
        workbook.Dispose();

        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        // download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
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
  