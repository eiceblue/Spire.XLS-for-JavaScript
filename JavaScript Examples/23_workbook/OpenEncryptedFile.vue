<template>
  <span>Click the following button to open an encrypted file</span>
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
        let excelFileName = "EncryptedFile.xlsx";
        await wasmModule.FetchFileToVFS(
          excelFileName,
          "",
          `${import.meta.env.BASE_URL}static/data/`
        );

        // Create string builder
        let builder = [];

        const passwords = ["password1", "password2", "password3", "1234"];
        for (let i = 0; i < passwords.length; i++) {
          try {
            // Create a workbook
            let workbook = wasmModule.Workbook.Create();

            // Open password
            workbook.OpenPassword = passwords[i];

            // Load the document
            workbook.LoadFromFile(excelFileName);

            builder.push(
              "Password = " +
                passwords[i] +
                " is correct. The encrypted Excel file opened successfully!"
            );
          } catch (e) {
            builder.push("Password = " + passwords[i] + " is not correct");
            builder.push("ErrorMessage = " + e.message); // Capture exception message
          }
        }

        let outputFileName = "OpenEncryptedFile_output.txt";
        wasmModule.FS.writeFile(outputFileName, builder.join("\n"));

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
  