<template>
  <span>Click the following button to open Excel files</span>
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
        let inputFile = "ExcelSample_N1.xlsx";
        let inputFile_97 = "ExcelSample97_N.xls";
        let inputFile_xml = "OfficeOpenXML_N.xml";
        let inputFile_csv = "CSVSample_N.csv";

        await wasmModule.FetchFileToVFS(
          inputFile,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );
        await wasmModule.FetchFileToVFS(
          inputFile_97,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );
        await wasmModule.FetchFileToVFS(
          inputFile_xml,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );
        await wasmModule.FetchFileToVFS(
          inputFile_csv,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create string builder
        let builder = [];

        // 1. Load file by file path
        // Create a workbook
        let workbook1 = wasmModule.Workbook.Create();
        // Load the document from disk
        workbook1.LoadFromFile({ fileName: inputFile });
        builder.push("Workbook opened using file path successfully!");

        // 2. Load file by file stream
        let stream = wasmModule.Stream.CreateByFile(inputFile.split("/").pop());
        // Create a workbook
        let workbook2 = wasmModule.Workbook.Create();
        // Load the document from stream
        workbook2.LoadFromStream(stream);
        builder.push("Workbook opened using file stream successfully!");
        stream.Close();

        // 3. Open Microsoft Excel 97 - 2003 file
        let wbExcel97 = wasmModule.Workbook.Create();
        wbExcel97.LoadFromFile({ fileName: inputFile_97 });
        builder.push("Microsoft Excel 97 - 2003 workbook opened successfully!");

        // 4. Open xml file
        let wbXML = wasmModule.Workbook.Create();
        wbXML.LoadFromXml(inputFile_xml);
        builder.push("XML file opened successfully!");

        // 5. Open csv file
        let wbCSV = wasmModule.Workbook.Create();
        wbCSV.LoadFromFile({
          fileName: inputFile_csv,
          separator: ",",
          row: 1,
          column: 1,
        });
        builder.push("CSV file opened successfully!");

        let outputFileName = "OpenFiles_output.txt";
        wasmModule.FS.writeFile(outputFileName, builder.join("\n"));
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
  