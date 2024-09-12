<template>
  <span
    >Click the following button to get excel properties and custom
    properties</span
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

        // Load the files
        let excelFileName = "WorksheetSample1.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const workbook = wasmModule.Workbook.Create();
        workbook.LoadFromFile(excelFileName);

        // Get the general excel properties
        let properties1 = workbook.DocumentProperties;
        let sb = [];
        sb.push("Excel Properties:");
        for (let i = 0; i < properties1.Count; i++) {
          let name = properties1.get(i).Name;
          let obj = properties1.get(i).Value;
          let t = properties1.get(i).PropertyType;
          let value = null;
          if (t === wasmModule.PropertyType.Double) {
            value = wasmModule.Double.Convert(obj).Value;
          } else if (t === wasmModule.PropertyType.DateTime) {
            value = wasmModule.DateTime.Convert(obj).ToString();
          } else if (t === wasmModule.PropertyType.Bool) {
            value = wasmModule.Boolean.Convert(obj).Value;
          } else if (
            t === wasmModule.PropertyType.Int ||
            t === wasmModule.PropertyType.Int32
          ) {
            value = wasmModule.Int32.Convert(obj).Value;
          } else {
            value = wasmModule.String.Convert(obj).Value;
          }
          sb.push(name + ": " + String(value));
        }
        sb.push("");

        // Get the custom properties
        let properties2 = workbook.CustomDocumentProperties;
        sb.push("Custom Properties:");
        for (let i = 0; i < properties2.Count; i++) {
          let name = properties2.get(i).Name;
          let t = properties2.get(i).PropertyType;
          let obj = properties2.get(i).Value;
          let value = null;
          if (t === wasmModule.PropertyType.Double) {
            value = wasmModule.Double.Convert(obj).Value;
          } else if (t === wasmModule.PropertyType.DateTime) {
            value = wasmModule.DateTime.Convert(obj).ToString();
          } else if (t === wasmModule.PropertyType.Bool) {
            value = wasmModule.Boolean.Convert(obj).Value;
          } else if (
            t === wasmModule.PropertyType.Int ||
            t === wasmModule.PropertyType.Int32
          ) {
            value = wasmModule.Int32.Convert(obj).Value;
          } else {
            value = wasmModule.String.Convert(obj).Value;
          }
          sb.push(name + ": " + String(value));
        }

        let outputFileName = "GetProperties_output.xlsx";
        wasmModule.FS.writeFile(outputFileName, sb.join("\n"));

        // Dispose of the workbook object to release resources
        book.Dispose();
        
        // Read the file from the virtual system and convert it to Blob
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: "text/plain",
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
  