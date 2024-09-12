<template>
    <span>The example demonstrates how to get the PivotTable refreshed information in worksheet.</span>
    <el-button @click="startProcessing">Start</el-button>
    <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">Click here to download the generated file</a>
</template>
<script>
import { ref } from 'vue';
export default {
    setup() {
        const downloadUrl = ref(null);
        const downloadName = ref('');

        const startProcessing = async () => {
            wasmModule = window.wasmModule;
            if (wasmModule) {
                // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
                await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

                // Input file
                let excelFileName = 'PivotTable.xlsx';
                await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

                // Create a new workbook
                const book = wasmModule.Workbook.Create();
                book.LoadFromFile({
                    fileName: excelFileName,
                    version: wasmModule.ExcelVersion.Version2010,
                });
                //Get first worksheet of the workbook
                let worksheet = book.Worksheets.get(0);

                //Get the first pivot table
                let pivotTable = worksheet.PivotTables.get(0);

                //Get the refreshed information
                let dateTime = pivotTable.Cache.RefreshDate;
                let refreshedBy = pivotTable.Cache.RefreshedBy;

                //Create StringBuilder to save
                let sb = [];

                //Set string format for displaying
                let result = 'Pivot table refreshed by:  ' + refreshedBy + '\r\nPivot table refreshed date: ' + dateTime.ToString();
                sb.push(result);
                // Define the output file name
                const outputFileName = 'GetPivotTableRefreshedInfo.txt';
                // Save result file
                wasmModule.FS.writeFile(outputFileName, sb.join('\n'));

                // Read the saved file and convert to a Blob object
                const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
                const modifiedFile = new Blob([modifiedFileArray], {
                    type: 'text/plain',
                });

                // Download the file
                downloadName.value  = outputFileName;
                downloadUrl.value = URL.createObjectURL(modifiedFile);

                // Clean up resources
                book.Dispose();
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
