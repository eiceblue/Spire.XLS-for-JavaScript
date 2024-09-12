<template>
    <span>The example demonstrates how to sort pivot table.</span>
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
                let excelFileName = 'SortPivotTable.xlsx';
                await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);
                // Create a new workbook
                const book = wasmModule.Workbook.Create();
                book.LoadFromFile({
                    fileName: excelFileName,
                    version: wasmModule.ExcelVersion.Version2010,
                });
                // Get the first worksheet
                let sheet = book.Worksheets.get(0);

                // Add an empty worksheet
                let sheet2 = book.CreateEmptySheet();
                sheet2.Name = 'Pivot Table';

                // Specify the data source
                let dataRange = sheet.Range.get('A1:C9');
                let cache = book.PivotCaches.Add({ range: dataRange });

                // Add PivotTable
                let pt = sheet2.PivotTables.Add('Pivot Table', sheet.Range.get('A1'), cache);

                // Configure the pivot table settings
                let r1 = pt.PivotFields.get_Item('No');
                r1.Axis = wasmModule.AxisTypes.Row;
                pt.Options.RowLayout = wasmModule.PivotTableLayoutType.Tabular;

                // Sort the "No" field in descending order
                r1.SortType = wasmModule.PivotFieldSortType.Descending;

                let r2 = pt.PivotFields.get_Item('Name');
                r2.Axis = wasmModule.AxisTypes.Row;
                // Add a data field to the pivot table
                pt.DataFields.Add(pt.PivotFields.get_Item('OnHand'), 'Sum of onHand', wasmModule.SubtotalTypes.None);
                // Set the pivot table style
                pt.BuiltInStyle = wasmModule.PivotBuiltInStyles.PivotStyleMedium12;
                // Define the output file name
                const outputFileName = 'SortPivotTable.xlsx';
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

