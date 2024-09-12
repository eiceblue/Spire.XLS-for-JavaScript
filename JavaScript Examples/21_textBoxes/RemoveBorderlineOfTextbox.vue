<template>
  <span
    >Click the following button to remove borderline of textbox in Excel
    chart</span
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
        workbook.Version = wasmModule.ExcelVersion.Version2013;

        //Create a new worksheet named "Remove Borderline" and add a chart to the worksheet
        let sheet = workbook.Worksheets.get(0);
        sheet.Name = "Remove Borderline";
        let chart = sheet.Charts.Add();

        //Create textbox1 in the chart and input text information
        let textbox1 = chart.TextBoxes.AddTextBox(50, 50, 100, 600);
        textbox1.Text = "The solution with borderline";

        //Create textbox2 in the chart, input text information and remove borderline
        let textbox2 = chart.TextBoxes.AddTextBox(1000, 50, 100, 600);
        textbox2.Text = "The solution without borderline";
        textbox2.Line.Weight = 0;

        let outputFileName = "RemoveBorderlineOfTextbox_output.xlsx";
        //Save the document
        workbook.SaveToFile({ fileName: outputFileName });

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
  