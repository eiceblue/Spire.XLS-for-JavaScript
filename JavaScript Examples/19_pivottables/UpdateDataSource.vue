<template>
    <span>
        The following example demonstrates how to update data source of pivot table.
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
                // Modify data of data source
                let data = book.Worksheets.get('Data');
                // Modify the data source by changing the value in cell A2 to "NewValue"
                data.Range.get('A2').Text = 'NewValue';
                // Modify the data source by changing the value in cell D2 to 28000
                data.Range.get('D2').NumberValue = 28000;

                // Get the sheet in which the pivot table is located
                let sheet = book.Worksheets.get({ sheetName: 'PivotTable' });
                // Get the first pivot table from the worksheet
                let pt = sheet.PivotTables.get(0);

                // Refresh and calculate
                pt.Cache.IsRefreshOnLoad = true;
                // Calculate and update the pivot table data
                pt.CalculateData();
                // Define the output file name
                const outputFileName = 'UpdateDataSource.xlsx';
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
