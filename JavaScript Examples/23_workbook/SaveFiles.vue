<template>
  <span>Click the following button to save Excel file to other formats</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>
  
  <script>
import { ref } from "vue";
import JSZip from "jszip";

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

        let outputDirectoryName = "outputFiles/";
        FS.mkdirTree(outputDirectoryName);
        // Load the files
        let excelFileName = "ExcelSample_N1.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        //Create a workbook and load a file
        const book = wasmModule.Workbook.Create();
        book.LoadFromFile(excelFileName);

        //Save to xls
        let output_xls = "SaveFiles_output.xls";
        book.SaveToFile({ fileName: outputDirectoryName + output_xls });

        //Save to xlsx
        let output_xlsx = "SaveFiles_output.xlsx";
        book.SaveToFile({ fileName: outputDirectoryName + output_xlsx });
        //Save to xlsb
        let output_xlsb = "SaveFiles_output.xlsb";
        book.SaveToFile({ fileName: outputDirectoryName + output_xlsb });
        //Save to ods
        let output_ods = "SaveFiles_output.ods";
        book.SaveToFile({ fileName: outputDirectoryName + output_ods });
        //Save to pdf
        let output_pdf = "SaveFiles_output.pdf";
        book.SaveToFile({ fileName: outputDirectoryName + output_pdf });
        //Save to xml
        let output_xml = "SaveFiles_output.xml";
        book.SaveToFile({ fileName: outputDirectoryName + output_xml });
        //Save to xps
        let output_xps = "SaveFiles_output.xps";
        book.SaveToFile({ fileName: outputDirectoryName + output_xps });

        // Dispose of the workbook object to release resources
        book.Dispose();

        const zip = new JSZip();
        let items = await FS.readdir(outputDirectoryName);
        items = items.filter((item) => item !== "." && item !== "..");
        for (const item of items) {
          const itemPath = `${outputDirectoryName}/${item}`;
          const fileData = await FS.readFile(itemPath);
          zip.file(item, fileData);
        }

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `outputFiles.zip`;
        downloadName.value = zipDownloadName;
        downloadUrl.value = zipDownloadUrl;
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
  