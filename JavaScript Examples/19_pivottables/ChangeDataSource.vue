<template>
    <span>
        The sample demonstrates how to modify the data source of the pivot
        table.
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
                let excelFileName = 'ChangeDataSource.xlsx';
                await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);

                // Create a new workbook
                const book = wasmModule.Workbook.Create();
                book.LoadFromFile({
                    fileName: excelFileName,
                    version: wasmModule.ExcelVersion.Version2010,
                });
                // Get the first worksheet
                let sheet = book.Worksheets.get(0);
                // Define the range of cells to be used as the new data source
                let Range = sheet.Range.get('A1:C15');
                // Get the first pivot table from the second worksheet
                let table = book.Worksheets.get(1).PivotTables.get(0);

                //Change data source
                table.ChangeDataSource(Range);
                table.Cache.IsRefreshOnLoad = false;

                // Define the output file name
                const outputFileName = 'ChangeDataSource.xlsx';
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


