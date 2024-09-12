<template>
  <span
    >Click the following button to access document properties by name and
    index</span
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
        let excelFileName = "AccessDocumentProperties.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile({ fileName: excelFileName });
        //Create string builder
        let builder = [];

        //Get all document properties
        let properties = book.CustomDocumentProperties;

        //Access document property by property name
        let property1 = properties.get({ strName: "Editor" });
        let obj = spirexls.String.Convert(property1.Value);
        builder.push(`${property1.Name} ${obj.Value}`);

        //Access document property by property index
        let property2 = properties.get({ iIndex: 0 });
        let obj2 = spirexls.String.Convert(property2.Value).Value;
        builder.push(`${property2.Name} ${obj2}`);

        let outputFileName = "AccessDocumentProperties_output.txt";
        wasmModule.FS.writeFile(outputFileName, builder.join("\n"));

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
