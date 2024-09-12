<template>
    <span>The following example demonstrates how to set the format of pivot field.</span>
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
                // Access the first pivot table in the worksheet
                let pt = sheet.PivotTables.get(0);
                // Access the first pivot field in the pivot table
                let pf = pt.PivotFields.get(0);

                //Setting the field auto sort ascend.
                pf.SortType = wasmModule.PivotFieldSortType.Ascending;

                //Setting Subtotal auto show.
                pf.SubtotalTop = true;

                //Setting Subtotal as Count type
                pf.Subtotals = wasmModule.SubtotalTypes.Count;

                //Setting the field auto show.
                pf.IsAutoShow = true;
                // Define the output file name
                const outputFileName = 'SetPivotFieldFormat.xlsx';
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
