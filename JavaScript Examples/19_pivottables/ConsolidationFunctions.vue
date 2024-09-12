<template>
    <span>
        The sample demonstrates how to use consolidation functions(e.g.Average, Max)
        of data field of pivot table.
    </span>
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
                let excelFileName = 'PivotTableExample.xlsx';
                await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

                // Create a new workbook
                const book = wasmModule.Workbook.Create();
                book.LoadFromFile({
                    fileName: excelFileName,
                    version: wasmModule.ExcelVersion.Version2010,
                });
                //Get the sheet in which the pivot table is located
                let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
                let pt = sheet.PivotTables.get(0);
                //Apply Average consolidation function to first data field
                pt.DataFields.get(0).Subtotal = wasmModule.SubtotalTypes.Average;
                //Apply Max consolidation function to second data field
                pt.DataFields.get(1).Subtotal = wasmModule.SubtotalTypes.Max;
                // Calculate the pivot table data
                pt.CalculateData();
                // Define the output file name
                const outputFileName = 'ConsolidationFunctions.xlsx';
                // Save the workbook to the specified path
                book.SaveToFile({
                    fileName: outputFileName,
                    version: wasmModule.ExcelVersion.Version2010,
                });

                // Read the saved file and convert to a Blob object
                const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
                const modifiedFile = new Blob([modifiedFileArray], {
                    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
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

