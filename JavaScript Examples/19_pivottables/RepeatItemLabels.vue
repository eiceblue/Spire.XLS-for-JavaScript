<template>
    <span>The sample demonstrates how to repeat item labels.</span>
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
                let excelFileName = 'RepeatItemLabelsExample.xlsx';
                await wasmModule.FetchFileToVFS(excelFileName, '', `${import.meta.env.BASE_URL}static/data/`);
                // Create a new workbook
                const book = wasmModule.Workbook.Create();
                book.LoadFromFile({
                    fileName: excelFileName,
                    version: wasmModule.ExcelVersion.Version2010,
                });
                //Get the first worksheet
                let sheet = book.Worksheets.get(0);
                //Add an empty worksheet
                let sheet2 = book.CreateEmptySheet();
                //Add PivotTable
                sheet2.Name = 'Pivot Table';
                // Define the data range for the pivot table
                let dataRange = sheet.Range.get('A1:D9');
                // Create a pivot cache using the data range
                let cache = book.PivotCaches.Add({ range: dataRange });
                // Add a pivot table to the pivot sheet using the pivot cache
                let pt = sheet2.PivotTables.Add('Pivot Table', sheet.Range.get('A1'), cache);
                // Set the VendorNo field as a row field and specify its header caption
                let r1 = pt.PivotFields.get_Item('VendorNo');
                r1.Axis = wasmModule.AxisTypes.Row;
                pt.Options.RowHeaderCaption = 'VendorNo';
                r1.Subtotals = wasmModule.SubtotalTypes.None;
                // Enable repeating item labels for the VendorNo field
                r1.RepeatItemLabels = true;
                // Enable repeating item labels for the OnHand field
                pt.PivotFields.get_Item('OnHand').RepeatItemLabels = true;
                 // Set the row layout type to tabular
                pt.Options.RowLayout = wasmModule.PivotTableLayoutType.Tabular;
                // Set the Desc field as an additional row field
                let r2 = pt.PivotFields.get_Item('Desc');
                r2.Axis = wasmModule.AxisTypes.Row;
                // Add the OnHand field as a data field with the label "Sum of onHand"
                pt.DataFields.Add(pt.PivotFields.get_Item('OnHand'), 'Sum of onHand', wasmModule.SubtotalTypes.None);
                // Set the built-in style for the pivot table appearance
                pt.BuiltInStyle = wasmModule.PivotBuiltInStyles.PivotStyleMedium12;

                // Define the output file name
                const outputFileName = 'RepeatItemLabels.xlsx';
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