<template>
  <span>Click the following button to convert workbook to HTML</span>
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
        await wasmModule.FetchFileToVFS(
          "ARIALUNI.TTF",
          "/Library/Fonts/",
          `${import.meta.env.BASE_URL}static/font/`
        );
        let outputDirectoryName = "WorkbookToHTMLFolder/";
        FS.mkdirTree(outputDirectoryName);
        let inputFileName = "WorkbookToHTML.xlsx";
        await wasmModule.FetchFileToVFS(
          inputFileName,
          outputDirectoryName,
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create a new workbook
        const workbook = wasmModule.Workbook.Create();
        // Load an existing Excel document
        workbook.LoadFromFile({
          fileName: outputDirectoryName + inputFileName,
        });

        const outputFileName = "WorkbookToHTML-out.html";
        // Save to HTML
        workbook.SaveToHtml({
          fileName: outputDirectoryName + outputFileName,
        });
        
        // Dispose of the workbook object to release resources
        workbook.Dispose();

        const zip = new JSZip();
        const addFilesToZip = async (folderPath, zipFolder) => {
          let items = await FS.readdir(folderPath);
          items = items.filter((item) => item !== "." && item !== "..");
          for (const item of items) {
            const itemPath = `${folderPath}/${item}`;
            try {
              const fileData = await FS.readFile(itemPath);
              zipFolder.file(item, fileData);
            } catch (error) {
              const zipSubFolder = zipFolder.folder(item);
              await addFilesToZip(itemPath, zipSubFolder);
            }
          }
        };

        await addFilesToZip(outputDirectoryName, zip);

        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `WorkbookToHTML.zip`;
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
