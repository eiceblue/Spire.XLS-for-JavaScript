<template>
  <div>
    <el-tree
      style="max-width: 600px"
      :data="data"
      :props="defaultProps"
      accordion
      @node-click="handleNodeClick"
    />
  </div>
</template>
<script setup lang="ts">
interface Tree {
  label: string;
  children?: Tree[];
  $treeNodeId?: string | number;
  [key: string]: any;
}
const emits = defineEmits(["item-id"]);
const handleNodeClick = (data: Tree) => {
  if ((data as any).id !== undefined) {
    emits("item-id", (data as any).id);
  }
};
const data: Tree[] = [
  {
    label: "Spire.XLS",
    children: [
      // 01_QuickGuide
      {
        label: "01_QuickGuide",
        children: [
          {
            label: "CreateAnExcelWithFiveSheets",
            id: "1-1",
          },
          {
            label: "CreateAnExcelWithOneSheet",
            id: "1-2",
          },
          {
            label: "CreateFiftyExcelFiles",
            id: "1-3",
          },
          {
            label: "HelloWorld",
            id: "1-4",
          },
          {
            label: "OpenExistingFile",
            id: "1-5",
          },
        ],
      },

      // 01_Data
      {
        label: "02_Data",
        children: [
          {
            label: "AddLabelControl",
            id: "2-1",
          },
          {
            label: "AddListBoxControl",
            id: "2-2",
          },
          {
            label: "AddScrollBarControl",
            id: "2-3",
          },
          {
            label: "AddTableWithFilter",
            id: "2-4",
          },
          {
            label: "AddTotalRowToTable",
            id: "2-5",
          },
          {
            label: "ApplySubscriptAndSuperscript",
            id: "2-6",
          },
          {
            label: "CloneExcelFontStyle",
            id: "2-7",
          },
          {
            label: "CopyCellsRange",
            id: "2-8",
          },
          {
            label: "CopyDataWithStyle",
            id: "2-9",
          },
          {
            label: "CopyOnlyFormulaValue",
            id: "2-10",
          },
          {
            label: "CreateNestedGroup",
            id: "2-11",
          },
          {
            label: "CreateTable",
            id: "2-12",
          },
          {
            label: "DataSorting",
            id: "2-13",
          },
          {
            label: "DataValidation",
            id: "2-14",
          },
          {
            label: "ExpandAndCollapseGroups",
            id: "2-15",
          },
          {
            label: "FindAndReplaceData",
            id: "2-16",
          },
          {
            label: "FindDataInSpecificRange",
            id: "2-17",
          },
          {
            label: "FindStringAndNumber",
            id: "2-18",
          },
          {
            label: "FormatTable",
            id: "2-19",
          },
          {
            label: "InsertControls",
            id: "2-20",
          },
          {
            label: "InsertHtmlStringIntoCell",
            id: "2-21",
          },
          {
            label: "ReplaceAndHighlight",
            id: "2-22",
          },
          {
            label: "RetrieveAndExtractData",
            id: "2-23",
          },
          {
            label: "SplitDataIntoMultipleColumns",
            id: "2-24",
          },
          {
            label: "Subtotal",
            id: "2-25",
          },
          {
            label: "WriteRichText",
            id: "2-26",
          },
        ],
      },
      // 03_Cells
      {
        label: "03_Cells",
        children: [
          {
            label: "AccessCell",
            id: "3-1",
          },
          {
            label: "ApplyMultipleFontsInSingleCell",
            id: "3-2",
          },
          {
            label: "AutoFitBasedOnCellValue",
            id: "3-3",
          },
          {
            label: "ConvertTextToNubmer",
            id: "3-4",
          },
          {
            label: "CopyCellFormat",
            id: "3-5",
          },
          {
            label: "CountNumberOfCells",
            id: "3-6",
          },
          {
            label: "CutCellsToOtherPosition",
            id: "3-7",
          },
          {
            label: "DetectMergedCells",
            id: "3-8",
          },
          {
            label: "DuplicateCellRange",
            id: "3-9",
          },
          {
            label: "EmptyCell",
            id: "3-10",
          },
          {
            label: "FilterCellsByCellColor",
            id: "3-11",
          },
          {
            label: "FindCellsWithStyleName",
            id: "3-12",
          },
          {
            label: "FindFormulaCells",
            id: "3-13",
          },
          {
            label: "GetCellAddress",
            id: "3-14",
          },
          {
            label: "GetCellDataType",
            id: "3-15",
          },
          {
            label: "GetCellDisplayedText",
            id: "3-16",
          },
          {
            label: "GetCellValueByCellName",
            id: "3-17",
          },
          {
            label: "GetIntersectionOfTwoRanges",
            id: "3-18",
          },
          {
            label: "HideCellContent",
            id: "3-19",
          },
          {
            label: "MergeCells",
            id: "3-20",
          },
          {
            label: "SetCellFillPattern",
            id: "3-21",
          },
          {
            label: "SetDBNumFormatting",
            id: "3-22",
          },
          {
            label: "ShrinkTextToFitInACell",
            id: "3-23",
          },
          {
            label: "TraverseCellsValue",
            id: "3-24",
          },
          {
            label: "UngroupExcelCells",
            id: "3-25",
          },
          {
            label: "UnmergeCells",
            id: "3-26",
          },
          {
            label: "UseExplicitLineBreaks",
            id: "3-27",
          },
          {
            label: "WrapOrUnwrapTextInCells",
            id: "3-28",
          },
        ],
      },
      // 04_RowsColumns
      {
        label: "04_RowsColumns",
        children: [
          {
            label: "AutoFitColumnInRange",
            id: "4-1",
          },
          {
            label: "AutoFitRowInRange",
            id: "4-2",
          },
          {
            label: "CheckAutoFitRowOrColumn",
            id: "4-3",
          },
          {
            label: "CheckRowOrColumnIsHidden",
            id: "4-4",
          },
          {
            label: "CopyColumns",
            id: "4-5",
          },
          {
            label: "CopyRows",
            id: "4-6",
          },
          {
            label: "CopySingleColumnAndRow",
            id: "4-7",
          },
          {
            label: "CopyWithOptions",
            id: "4-8",
          },
          {
            label: "DeleteBlankRowsAndColumns",
            id: "4-9",
          },
          {
            label: "DeleteMultipleRowsAndColumns",
            id: "4-10",
          },
          {
            label: "GetDefaultRowAndColumnCount",
            id: "4-11",
          },
          {
            label: "GroupRowsAndColumns",
            id: "4-12",
          },
          {
            label: "HideOrShowRowColumnHeaders",
            id: "4-13",
          },
          {
            label: "HideRowsAndColumns",
            id: "4-14",
          },
          {
            label: "InsertRowsAndColumns",
            id: "4-15",
          },
          {
            label: "RemoveRowBasedOnKeyword",
            id: "4-16",
          },
          {
            label: "SetColumnWithInPixels",
            id: "4-17",
          },
          {
            label: "SetDefaultColumnWidth",
            id: "4-18",
          },
          {
            label: "SetDefaultRowAndColumnStyle",
            id: "4-19",
          },
          {
            label: "SetDefaultRowHeight",
            id: "4-20",
          },
          {
            label: "SetHeightAndWidth",
            id: "4-21",
          },
          {
            label: "SetSummaryColumnDirection",
            id: "4-22",
          },
          {
            label: "SetSummaryRowDirection",
            id: "4-23",
          },
          {
            label: "UnhideRowsAndColumns",
            id: "4-24",
          },
        ],
      },
      //05_Images
      {
        label: "05_Images",
        children: [
          {
            label: "AlignPictureWithinCell",
            id: "5-1",
          },
          {
            label: "CompressPictures",
            id: "5-2",
          },
          {
            label: "CopyPicture",
            id: "5-3",
          },
          {
            label: "CroppedPositionOfPicture",
            id: "5-4",
          },
          {
            label: "DeleteAllImages",
            id: "5-5",
          },
          {
            label: "InsertExcelBackgroundImage",
            id: "5-6",
          },
          {
            label: "LocateImages",
            id: "5-7",
          },
          {
            label: "PictureOffset",
            id: "5-8",
          },
          {
            label: "PictureRefRange",
            id: "5-9",
          },
          {
            label: "ReadImages",
            id: "5-10",
          },
          {
            label: "RemovePictureBorder",
            id: "5-11",
          },
          {
            label: "ResetSizeAndPositionForImage",
            id: "5-12",
          },
          {
            label: "SetImageOffsetOfChart",
            id: "5-13",
          },
          {
            label: "WriteImages",
            id: "5-14",
          },
        ],
      },
      //06_Comments
      {
        label: "06_Comments",
        children: [
          {
            label: "AddCommentWithAuthor",
            id: "6-1",
          },
          {
            label: "AddCommentWithPicture",
            id: "6-2",
          },
          {
            label: "EditExcelComment",
            id: "6-3",
          },
          {
            label: "HideOrShowComment",
            id: "6-4",
          },
          {
            label: "ReadComment",
            id: "6-5",
          },
          {
            label: "RemoveComment",
            id: "6-6",
          },
          {
            label: "SetCommentFillColor",
            id: "6-7",
          },
          {
            label: "SetCommentTextRotation",
            id: "6-8",
          },
          {
            label: "SetPositionAndAlignment",
            id: "6-9",
          },
          {
            label: "WriteComment",
            id: "6-10",
          },
        ],
      },
      //07_Conversion
      {
        label: "07_Conversion",
        children: [
          {
            label: "ChartSheetToSVG",
            id: "7-1",
          },
          {
            label: "CSVToExcel",
            id: "7-2",
          },
          {
            label: "CSVToPDF",
            id: "7-3",
          },
          {
            label: "EachWorksheetToDifferentPDF",
            id: "7-4",
          },
          {
            label: "FitWidthWhenConvertToPDF",
            id: "7-5",
          },
          {
            label: "HtmlToExcel",
            id: "7-6",
          },
          {
            label: "LoadSaveEtAndETT",
            id: "7-7",
          },
          {
            label: "OfficeOpenXMLToExcel",
            id: "7-8",
          },
          {
            label: "SelectedRangeToPDF",
            id: "7-9",
          },
          {
            label: "SheetToImage",
            id: "7-10",
          },
          {
            label: "SpecificCellsToImage",
            id: "7-11",
          },
          {
            label: "SpecifyFontDirectory",
            id: "7-12",
          },
          {
            label: "ToCSV",
            id: "7-13",
          },
          {
            label: "ToCSVWithFilteredValue",
            id: "7-14",
          },
          {
            label: "ToHtml",
            id: "7-15",
          },
          {
            label: "ToHtmlStream",
            id: "7-16",
          },
          {
            label: "ToImageWithoutWhiteSpace",
            id: "7-17",
          },
          {
            label: "ToODS",
            id: "7-18",
          },
          {
            label: "ToOFD",
            id: "7-19",
          },
          {
            label: "ToOfficeOpenXML",
            id: "7-20",
          },
          {
            label: "ToPDF",
            id: "7-21",
          },
          {
            label: "ToPDFA1B",
            id: "7-22",
          },
          {
            label: "ToPdfSimply",
            id: "7-23",
          },
          {
            label: "ToPdfWithChangePageSize",
            id: "7-24",
          },
          {
            label: "ToPostScript",
            id: "7-25",
          },
          {
            label: "ToSVG",
            id: "7-26",
          },
          {
            label: "ToText",
            id: "7-27",
          },
          {
            label: "ToXPS",
            id: "7-28",
          },
          {
            label: "WorkbookToHTML",
            id: "7-29",
          },
          {
            label: "XLSToXLSM",
            id: "7-30",
          },
        ],
      },
      //08_FilteringAndValidation
      {
        label: "08_FilteringAndValidation",
        children: [
          {
            label: "AutofilterBlank",
            id: "8-1",
          },
          {
            label: "AutofilterNonBlank",
            id: "8-2",
          },
          {
            label: "CreateFilter",
            id: "8-3",
          },
          {
            label: "FilterCellsByString",
            id: "8-4",
          },
          {
            label: "GetSettingsOfDataValidation",
            id: "8-5",
          },
          {
            label: "ListDataValidation",
            id: "8-6",
          },
          {
            label: "RemoveAutoFilters",
            id: "8-7",
          },
          {
            label: "RemoveDataValidation",
            id: "8-8",
          },
          {
            label: "SetDataValidationOnSeparateSheet",
            id: "8-9",
          },
          {
            label: "TimeDataValidation",
            id: "8-10",
          },
          {
            label: "VerifyDataByValidation",
            id: "8-11",
          },
          {
            label: "WholeNumberDataValidation",
            id: "8-12",
          },
        ],
      },
      // 09_Chart
      {
        label: "09_Chart",
        children: [
          {
            label: "AddDataTable",
            id: "9-1",
          },
          {
            label: "AddErrorBars",
            id: "9-2",
          },
          {
            label: "AddPictureInChart",
            id: "9-3",
          },
          {
            label: "AddTextBox",
            id: "9-4",
          },
          {
            label: "AddTrendline",
            id: "9-5",
          },
          {
            label: "AjustBarSpace",
            id: "9-6",
          },
          {
            label: "ApplySoftEdgesEffect",
            id: "9-7",
          },
          {
            label: "ChangeChartSizeAndPosition",
            id: "9-8",
          },
          {
            label: "ChangeDataLabel",
            id: "9-9",
          },
          {
            label: "ChangeDataRange",
            id: "9-10",
          },
          {
            label: "ChangeMajorGridlines",
            id: "9-11",
          },
          {
            label: "ChangeSeriesColor",
            id: "9-12",
          },
          {
            label: "ChartAxisTitle",
            id: "9-13",
          },
          {
            label: "ChartToImage",
            id: "9-14",
          },
          {
            label: "ClusteredBar",
            id: "9-15",
          },
          {
            label: "ClusteredBar_3D",
            id: "9-16",
          },
          {
            label: "ClusteredColumn",
            id: "9-17",
          },
          {
            label: "CreateBoxAndWhiskerChart",
            id: "9-18",
          },
          {
            label: "CreateBubbleChart",
            id: "9-19",
          },
          {
            label: "CreateChartBasedOnPivotTable",
            id: "9-20",
          },
          {
            label: "CreateChartWithoutRangeData",
            id: "9-21",
          },
          {
            label: "CreateCustomChart",
            id: "9-22",
          },
          {
            label: "CreateDoughnutChart",
            id: "9-23",
          },
          {
            label: "CreateFunnelChart",
            id: "9-24",
          },
          {
            label: "CreateHistogramChart",
            id: "9-25",
          },
          {
            label: "CreateMultiLevelChart",
            id: "9-26",
          },
          {
            label: "CreateParetoChart",
            id: "9-27",
          },
          {
            label: "CreatePivotChart",
            id: "9-28",
          },
          {
            label: "CreateRadarChart",
            id: "9-29",
          },
          {
            label: "CreateSunBurstChart",
            id: "9-30",
          },
          {
            label: "CreateTreeMapChart",
            id: "9-31",
          },
          {
            label: "CreateWaterfallChart",
            id: "9-32",
          },
          {
            label: "CustomDataMarker",
            id: "9-33",
          },
          {
            label: "DataCallout",
            id: "9-34",
          },
          {
            label: "DeleteLegend",
            id: "9-35",
          },
          {
            label: "DiscontinuousData",
            id: "9-36",
          },
          {
            label: "EditLineChart",
            id: "9-37",
          },
          {
            label: "ExplodedDoughnut",
            id: "9-38",
          },
          {
            label: "ExtractTrendline",
            id: "9-39",
          },
          {
            label: "FillChartElementWithPicture",
            id: "9-40",
          },
          {
            label: "FormatAxis",
            id: "9-41",
          },
          {
            label: "GaugeChart",
            id: "9-42",
          },
          {
            label: "GetCategoryLabels",
            id: "9-43",
          },
          {
            label: "GetChartDataPointValues",
            id: "9-44",
          },
          {
            label: "GetWorksheetOfChart",
            id: "9-45",
          },
          {
            label: "HideMajorGridlines",
            id: "9-46",
          },
          {
            label: "Line",
            id: "9-47",
          },
          {
            label: "Pie",
            id: "9-48",
          },
          {
            label: "PyramidColumn",
            id: "9-49",
          },
          {
            label: "RemoveChart",
            id: "9-50",
          },
          {
            label: "ResizeAndMoveChart",
            id: "9-51",
          },
          {
            label: "RichTextForDataLabel",
            id: "9-52",
          },
          {
            label: "Rotate3DChart",
            id: "9-53",
          },
          {
            label: "ScatterChart",
            id: "9-54",
          },
          {
            label: "SetAndFormatDataLabel",
            id: "9-55",
          },
          {
            label: "SetBorderColorAndStyle",
            id: "9-56",
          },
          {
            label: "SetBorderWidthOfMarker",
            id: "9-57",
          },
          {
            label: "SetChartBackgroundColor",
            id: "9-58",
          },
          {
            label: "SetColorForChartArea",
            id: "9-59",
          },
          {
            label: "SetFont",
            id: "9-60",
          },
          {
            label: "SetFontForLegendAndDataTable",
            id: "9-61",
          },
          {
            label: "SetFontForTitleAndAxis",
            id: "9-62",
          },
          {
            label: "SetLegendBackgroundColor",
            id: "9-63",
          },
          {
            label: "SetNumberFormatOfTrendline",
            id: "9-64",
          },
          {
            label: "ShowLeaderLine",
            id: "9-65",
          },
          {
            label: "ShowLeaderLine",
            id: "9-66",
          },
          {
            label: "StackedColumn",
            id: "9-67",
          },
        ],
      },
      // 10_Shape
      {
        label: "10_Shapes",
        children: [
          {
            label: "AddArrowLineToExcelFile",
            id: "10-1",
          },
          {
            label: "AddLineShape",
            id: "10-2",
          },
          {
            label: "AddOvalShape",
            id: "10-3",
          },
          {
            label: "AddRectangleShape",
            id: "10-4",
          },
          {
            label: "AddSpinnerControl",
            id: "10-5",
          },
          {
            label: "AdjustArrowPolylinePosition",
            id: "10-6",
          },
          {
            label: "CopyShapes",
            id: "10-7",
          },
          {
            label: "DeleteAllShapes",
            id: "10-8",
          },
          {
            label: "DeleteParticularShape",
            id: "10-9",
          },
          {
            label: "DrawOneLineThroughTwoPoints",
            id: "10-10",
          },
          {
            label: "ExtractTextImageFromShape",
            id: "10-11",
          },
          {
            label: "GetShapeLinkedCellRange",
            id: "10-12",
          },
          {
            label: "HideOrUnhideShape",
            id: "10-13",
          },
          {
            label: "InsertShapesToExcelSheet",
            id: "10-14",
          },
          {
            label: "ModifyShadowStyleForShape",
            id: "10-15",
          },
          {
            label: "SetShadowStyleForShape",
            id: "10-16",
          },
          {
            label: "SetShapeOrder",
            id: "10-17",
          },
          {
            label: "ShapeToImage",
            id: "10-18",
          },
          {
            label: "TillPicAsTextureInShape",
            id: "10-19",
          },
        ],
      },
      //11_Formatting
      {
        label: "11_Formatting",
        children: [
          {
            label: "ApplyBuiltInStyles",
            id: "11-1",
          },
          {
            label: "ApplyColorScalesToDataRange",
            id: "11-2",
          },
          {
            label: "ApplyConditionalFormatting",
            id: "11-3",
          },
          {
            label: "ApplyDataBarsToCellRange",
            id: "11-4",
          },
          {
            label: "ApplyGradientFillEffects",
            id: "11-5",
          },
          {
            label: "ApplyIconSetsToCellRange",
            id: "11-6",
          },
          {
            label: "ColorsAndPalette",
            id: "11-7",
          },
          {
            label: "ConditionalFormatRuntime",
            id: "11-8",
          },
          {
            label: "ConditionallyFormatDate",
            id: "11-9",
          },
          {
            label: "CreateFormulaConditionalFormat",
            id: "11-10",
          },
          {
            label: "FontStyles",
            id: "11-11",
          },
          {
            label: "ForegroundAndBackground",
            id: "11-12",
          },
          {
            label: "FormatAColumn",
            id: "11-13",
          },
          {
            label: "FormatARow",
            id: "11-14",
          },
          {
            label: "FormatCellsWithStyle",
            id: "11-15",
          },
          {
            label: "GetColorArgbData",
            id: "11-16",
          },
          {
            label: "GetStyleSetStyle",
            id: "11-17",
          },
          {
            label: "HighlightAverageValues",
            id: "11-18",
          },
          {
            label: "HighlightDuplicateUniqueValues",
            id: "11-19",
          },
          {
            label: "HighlightRankedValues",
            id: "11-20",
          },
          {
            label: "Indentation",
            id: "11-21",
          },
          {
            label: "Interior",
            id: "11-22",
          },
          {
            label: "MakeCellActive",
            id: "11-23",
          },
          {
            label: "NumberStyles",
            id: "11-24",
          },
          {
            label: "SetBorder",
            id: "11-25",
          },
          {
            label: "SetBorderToDataBar",
            id: "11-26",
          },
          {
            label: "SetConditionalFormatFormula",
            id: "11-27",
          },
          {
            label: "SetRowColorByConditionalFormat",
            id: "11-28",
          },
          {
            label: "SetTrafficLightsIcons",
            id: "11-29",
          },
          {
            label: "SimpleConditionalFormatting",
            id: "11-30",
          },
          {
            label: "TextAlign",
            id: "11-31",
          },
          {
            label: "TextDirection",
            id: "11-32",
          },
          {
            label: "UsePredefinedStyles",
            id: "11-33",
          },
          {
            label: "UsingStyleObject",
            id: "11-34",
          },
          {
            label: "VariousConditionalFormatting",
            id: "11-35",
          },
        ],
      },
      //12_Formulas
      {
        label: "12_Formulas",
        children: [
          {
            label: "InsertFormulaWithNamedRange",
            id: "12-1",
          },
          {
            label: "ReadFormulas",
            id: "12-2",
          },
          {
            label: "RegisterAddInFunction",
            id: "12-3",
          },
          {
            label: "RemoveFormulasButKeepValues",
            id: "12-4",
          },
          {
            label: "SubTotalFormula",
            id: "12-5",
          },
          {
            label: "UseArrayFormulas",
            id: "12-6",
          },
          {
            label: "UseArrayR1C1Formula",
            id: "12-7",
          },
          {
            label: "UseR1C1Formula",
            id: "12-8",
          },
          {
            label: "WriteFormulas",
            id: "12-9",
          },
        ],
      },
      //13_HeaderFooter
      {
        label: "13_HeaderFooter",
        children: [
          {
            label: "ChangeFontAndSize",
            id: "13-1",
          },
          {
            label: "DifferentHeaderFooter",
            id: "13-2",
          },
          {
            label: "DifferentHeaderFooterOnFirstPage",
            id: "13-3",
          },
          {
            label: "ImageHeaderFooter",
            id: "13-4",
          },
          {
            label: "SetHeaderFooter",
            id: "13-5",
          },
        ],
      },
      //14_Hyperlinks
      {
        label: "14_Hyperlinks",
        children: [
          {
            label: "AddHyperlinkToText",
            id: "14-1",
          },
          {
            label: "AddImageHyperlink",
            id: "14-2",
          },
          {
            label: "GetHyperLinkType",
            id: "14-3",
          },
          {
            label: "LinkToOtherSheetCell",
            id: "14-4",
          },
          {
            label: "ModifyHyperlink",
            id: "14-5",
          },
          {
            label: "ReadHyperlinks",
            id: "14-6",
          },
          {
            label: "RemoveHyperlinks",
            id: "14-7",
          },
          {
            label: "RetrieveExternalFileHyperlinks",
            id: "14-8",
          },
          {
            label: "WriteHyperlinks",
            id: "14-9",
          },
          {
            label: "LinkToExternalFile",
            id: "14-10",
          },
        ],
      },
      //15_MarkerDesigner
      {
        label: "15_MarkerDesigner",
        children: [
          {
            label: "AddVariableArray",
            id: "15-1",
          },
        ],
      },
      //16_Namedranges
      {
        label: "16_Namedranges",
        children: [
          {
            label: "FormatNamedRangeCells",
            id: "16-1",
          },
          {
            label: "GetAllNamedRange",
            id: "16-2",
          },
          {
            label: "GetNamedRangeAddress",
            id: "16-3",
          },
          {
            label: "GetSpecificNamedRange",
            id: "16-4",
          },
          {
            label: "MergeNamedRangeCells",
            id: "16-5",
          },
          {
            label: "NamedRanges",
            id: "16-6",
          },
          {
            label: "RemoveNamedRange",
            id: "16-7",
          },
          {
            label: "RenameNamedRange",
            id: "16-8",
          },
          {
            label: "ScopedNamedRange",
            id: "16-9",
          },
          {
            label: "SetFormulaWithNamedRange",
            id: "16-10",
          },
        ],
      },
      //17_Oleobjects
      {
        label: "17_Oleobjects",
        children: [
          {
            label: "ExtractOLEObjects",
            id: "17-1",
          },
          {
            label: "InsertOLEObjects",
            id: "17-2",
          },
          {
            label: "InsertWavFileOLEObject",
            id: "17-3",
          },
        ],
      },
      //18_Pagesetup
      {
        label: "18_Pagesetup",
        children: [
          {
            label: "GetExcelPaperDimensions",
            id: "18-1",
          },
          {
            label: "GetExcelVersion",
            id: "18-2",
          },
          {
            label: "SetExcelPageOrderType",
            id: "18-3",
          },
          {
            label: "SetExcelPaperSize",
            id: "18-4",
          },
          {
            label: "SetFirstPageNumber",
            id: "18-5",
          },
          {
            label: "SetHeaderAndFooterMargins",
            id: "18-6",
          },
          {
            label: "SetMarginsOfExcelSheet",
            id: "18-7",
          },
          {
            label: "SetOtherPrintingOptions",
            id: "18-8",
          },
          {
            label: "SetPageOrientation",
            id: "18-9",
          },
          {
            label: "SetPrintAreaOfXlsFile",
            id: "18-10",
          },
          {
            label: "SetPrintQualityOfXlsFile",
            id: "18-11",
          },
          {
            label: "SetPrintTitleOfXlsFile",
            id: "18-12",
          },
          {
            label: "SetSheetFitToPageProperty",
            id: "18-13",
          },
          {
            label: "SetXlsSheetCenterOnPage",
            id: "18-14",
          },
        ],
      },
      //19_Pivottables
      {
        label: "19_Pivottables",
        children: [
          {
            label: "ChangeDataSource",
            id: "19-1",
          },
          {
            label: "ClearPivotFields",
            id: "19-2",
          },
          {
            label: "ConsolidationFunctions",
            id: "19-3",
          },
          {
            label: "CreatePivotTable",
            id: "19-4",
          },
          {
            label: "DisablePivotTableRibbon",
            id: "19-5",
          },
          {
            label: "ExpandOrCollapseRows",
            id: "19-6",
          },
          {
            label: "FormatDataField",
            id: "19-7",
          },
          {
            label: "FormattingAppearance",
            id: "19-8",
          },
          {
            label: "GetPivotTableRefreshedInfo",
            id: "19-9",
          },
          {
            label: "RefreshPivotTable",
            id: "19-10",
          },
          {
            label: "RepeatItemLabels",
            id: "19-11",
          },
          {
            label: "SetFormatOptions",
            id: "19-12",
          },
          {
            label: "SetPivotFieldFormat",
            id: "19-13",
          },
          {
            label: "ShowDataFieldInRow",
            id: "19-14",
          },
          {
            label: "ShowSubTotals",
            id: "19-15",
          },
          {
            label: "SortPivotTable",
            id: "19-16",
          },
          {
            label: "UpdateDataSource",
            id: "19-17",
          },
        ],
      },
      //20_Security
      {
        label: "20_Security",
        children: [
          {
            label: "DetectProtection",
            id: "20-1",
          },
          {
            label: "HideFormulas",
            id: "20-2",
          },
          {
            label: "LockSpecificCellInNewExcel",
            id: "20-3",
          },
          {
            label: "LockSpecificColumnInNewExcel",
            id: "20-4",
          },
          {
            label: "LockSpecificRowInNewExcel",
            id: "20-5",
          },
          {
            label: "ProtectCell",
            id: "20-6",
          },
          {
            label: "ProtectWithEditableRange",
            id: "20-7",
          },
          {
            label: "ProtectWorkbook",
            id: "20-8",
          },
          {
            label: "RemoveAllDigitalSignatures",
            id: "20-9",
          },
          {
            label: "UnlockProtectSheet",
            id: "20-10",
          },
          {
            label: "UnlockSimpleSheet",
            id: "20-11",
          },
        ],
      },
      //21_TextBoxes
      {
        label: "21_TextBoxes",
        children: [
          {
            label: "ExtractTextFromATextbox",
            id: "21-1", 
          },
          {
            label: "GetTextBoxByName",
            id: "21-2",
          },
          {
            label: "ManipulateTextBox",
            id: "21-3",
          },
          {
            label: "RemoveBorderlineOfTextbox",
            id: "21-4",
          },
          {
            label: "ReplaceTextInTextBox",
            id: "21-5",
          },
          {
            label: "SetFontAndBackground",
            id: "21-6",
          },
          {
            label: "SetInternalMarginOfTextbox",
            id: "21-7",
          },
          {
            label: "TextBoxWithWrapText",
            id: "21-8",
          },
        ],
      },
      //22_Worksheets
      {
        label: "22_Worksheets",
        children: [
          {
            label: "ActivateWorksheet",
            id: "22-1", 
          },
          {
            label: "AddPageBreakInXlsFile",
            id: "22-2",
          },
          {
            label: "AddWorksheet",
            id: "22-3",
          },
          {
            label: "ApplyStyleToWorksheet",
            id: "22-4",
          },
          {
            label: "CheckDialogSheetInXlsFile",
            id: "22-5",
          },
          {
            label: "CopySheetToAnotherXlsFile",
            id: "22-6",
          },     
          {
            label: "CopySheetWithinWorkbook",
            id: "22-7",
          },
          {
            label: "CopyVisibleSheets",
            id: "22-8",
          },
          {
            label: "CopyWorksheet",
            id: "22-9",
          },
          {
            label: "DetectEmptyWorksheet",
            id: "22-10",
          },
          {
            label: "FillDataInWorksheet",
            id: "22-11",
          },
          {
            label: "FreezePanes",
            id: "22-12",
          },
          {
            label: "GetFreezePaneRange",
            id: "22-13",
          },
          {
            label: "GetListOfFontsUsed",
            id: "22-14",
          },
          {
            label: "GetPaperSize",
            id: "22-15",
          },
          {
            label: "GetWorksheetNames",
            id: "22-16",
          },
          {
            label: "HideOrShowWorksheet",
            id: "22-17",
          },
          {
            label: "HideTab",
            id: "22-18",
          },
          {
            label: "HideZeroValues",
            id: "22-19",
          },
          {
            label: "LinkToContentProperty",
            id: "22-20",
          },
          {
            label: "MoveWorksheet",
            id: "22-21",
          },
          {
            label: "PageBreakPreview",
            id: "22-22",
          },
          {
            label: "RemovePageBreak",
            id: "22-23",
          },
          {
            label: "RemoveWorksheet",
            id: "22-24",
          },
          {
            label: "SetPageBreak",
            id: "22-25",
          },
          {
            label: "SetTabColor",
            id: "22-26",
          },
          {
            label: "SetViewMode",
            id: "22-27",
          },
          {
            label: "ShowOrHideGridLine",
            id: "22-28",
          },
          {
            label: "ShowTab",
            id: "22-29",
          },
          {
            label: "SplitWorksheetIntoPanes",
            id: "22-30",
          },
          {
            label: "UnfreezeExcelPanes",
            id: "22-31",
          },
          {
            label: "VerifyProtectedWorksheet",
            id: "22-32",
          },
          {
            label: "ZoomFactor",
            id: "22-33",
          },
          
        ],
      },
      //23_Workbook
      {
        label: "23_Workbook",
        children: [
          {
            label: "AccessDocumentProperties",
            id: "23-1", 
          },
          {
            label: "AddCustomProperties",
            id: "23-2",
          },
          {
            label: "DecryptWorkbook",
            id: "23-3",
          },          
          {
            label: "DetectExcelVersion",
            id: "23-4",
          },
          {
            label: "DetectVBAMacros",
            id: "23-5",
          },
          {
            label: "EncryptWorkbook",
            id: "23-6",
          },
          {
            label: "GetProperties",
            id: "23-7",
          },
          {
            label: "HideWindow",
            id: "23-8",
          },
          {
            label: "LoadAndSaveFileWithMacro",
            id: "23-9",
          },
          {
            label: "MergeExcelFiles",
            id: "23-10",
          },
          {
            label: "OpenEncryptedFile",
            id: "23-11",
          },
          {
            label: "OpenFiles",
            id: "23-12",
          },
          {
            label: "ReadStream",
            id: "23-13",
          },
          {
            label: "RemoveCustomProperties",
            id: "23-14",
          },
          {
            label: "SaveFiles",
            id: "23-15",
          },
          {
            label: "SaveStream",
            id: "23-16",
          },
          {
            label: "SetExcelCalculationMode",
            id: "23-17",
          },
          {
            label: "SetMargins",
            id: "23-18",
          },
          {
            label: "SetTheme",
            id: "23-19",
          },
        ],
      },
      //24_TrackChanges
      {
        label: "24_TrackChanges",
        children: [
          {
            label: "AcceptOrRejectTrackedChanges",
            id: "24-1", 
          },
        ],
      },
    ],
  },
];
const defaultProps = {
  children: "children",
  label: "label",
};
</script>