<template>
    <span>The sample demonstrates how to set the format of pivot table.</span>
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
                //Set the PivotTable report is automatically formatted
                pt.Options.IsAutoFormat = true;

                //Setting the PivotTable report shows grand totals for rows.
                pt.ShowRowGrand = true;

                //Setting the PivotTable report shows grand totals for columns.
                pt.ShowColumnGrand = true;

                //Setting the PivotTable report displays a custom string in cells that contain null values.
                pt.DisplayNullString = true;
                pt.NullString = 'null';

                //Setting the PivotTable report's layout
                pt.PageFieldOrder = wasmModule.PagesOrderType.DownThenOver;

                // Define the output file name
                const outputFileName = 'SetFormatOptions.xlsx';
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
