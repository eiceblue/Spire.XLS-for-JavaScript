<template>
  <div class="worker_page">
    <div class="tree_item">
      <SpireTree @item-id="handleChildID" />
    </div>
    <div class="file_item">
      <div class="file-reader">
        <div v-if="filePageId == '0'" class="non-centered">
          <Description />
        </div>
        <!--01_quickguide -->
        <div v-if="filePageId == '1-1'">
          <CreateAnExcelWithFiveSheets />
        </div>
        <div v-if="filePageId == '1-2'">
          <CreateAnExcelWithOneSheet />
        </div>
        <div v-if="filePageId == '1-3'">
          <CreateFiftyExcelFiles />
        </div>
        <div v-if="filePageId == '1-4'">
          <HelloWorld />
        </div>
        <div v-if="filePageId == '1-5'">
          <OpenExistingFile />
        </div>
        <!--02_data -->
        <div v-if="filePageId == '2-1'">
          <AddLabelControl />
        </div>
        <div v-if="filePageId == '2-2'">
          <AddListBoxControl />
        </div>
        <div v-if="filePageId == '2-3'">
          <AddScrollBarControl />
        </div>
        <div v-if="filePageId == '2-4'">
          <AddTableWithFilter />
        </div>
        <div v-if="filePageId == '2-5'">
          <AddTotalRowToTable />
        </div>
        <div v-if="filePageId == '2-6'">
          <ApplySubscriptAndSuperscript />
        </div>
        <div v-if="filePageId == '2-7'">
          <CloneExcelFontStyle />
        </div>
        <div v-if="filePageId == '2-8'">
          <CopyCellsRange />
        </div>
        <div v-if="filePageId == '2-9'">
          <CopyDataWithStyle />
        </div>
        <div v-if="filePageId == '2-10'">
          <CopyOnlyFormulaValue />
        </div>
        <div v-if="filePageId == '2-11'">
          <CreateNestedGroup />
        </div>
        <div v-if="filePageId == '2-12'">
          <CreateTable />
        </div>
        <div v-if="filePageId == '2-13'">
          <DataSorting />
        </div>
        <div v-if="filePageId == '2-14'">
          <DataValidation />
        </div>
        <div v-if="filePageId == '2-15'">
          <ExpandAndCollapseGroups />
        </div>
        <div v-if="filePageId == '2-16'">
          <FindAndReplaceData />
        </div>
        <div v-if="filePageId == '2-17'">
          <FindDataInSpecificRange />
        </div>
        <div v-if="filePageId == '2-18'">
          <FindStringAndNumber />
        </div>
        <div v-if="filePageId == '2-19'">
          <FormatTable />
        </div>
        <div v-if="filePageId == '2-20'">
          <InsertControls />
        </div>
        <div v-if="filePageId == '2-21'">
          <InsertHtmlStringIntoCell />
        </div>
        <div v-if="filePageId == '2-22'">
          <ReplaceAndHighlight />
        </div>
        <div v-if="filePageId == '2-23'">
          <RetrieveAndExtractData />
        </div>
        <div v-if="filePageId == '2-24'">
          <SplitDataIntoMultipleColumns />
        </div>
        <div v-if="filePageId == '2-25'">
          <Subtotal />
        </div>
        <div v-if="filePageId == '2-26'">
          <WriteRichText />
        </div>
        <!--03_cells -->
        <div v-if="filePageId == '3-1'">
          <AccessCell />
        </div>
        <div v-if="filePageId == '3-2'">
          <ApplyMultipleFontsInSingleCell />
        </div>
        <div v-if="filePageId == '3-3'">
          <AutoFitBasedOnCellValue />
        </div>
        <div v-if="filePageId == '3-4'">
          <ConvertTextToNubmer />
        </div>
        <div v-if="filePageId == '3-5'">
          <CopyCellFormat />
        </div>
        <div v-if="filePageId == '3-6'">
          <CountNumberOfCells />
        </div>
        <div v-if="filePageId == '3-7'">
          <CutCellsToOtherPosition />
        </div>
        <div v-if="filePageId == '3-8'">
          <DetectMergedCells />
        </div>
        <div v-if="filePageId == '3-9'">
          <DuplicateCellRange />
        </div>
        <div v-if="filePageId == '3-10'">
          <EmptyCell />
        </div>
        <div v-if="filePageId == '3-11'">
          <FilterCellsByCellColor />
        </div>
        <div v-if="filePageId == '3-12'">
          <FindCellsWithStyleName />
        </div>
        <div v-if="filePageId == '3-13'">
          <FindFormulaCells />
        </div>
        <div v-if="filePageId == '3-14'">
          <GetCellAddress />
        </div>
        <div v-if="filePageId == '3-15'">
          <GetCellDataType />
        </div>
        <div v-if="filePageId == '3-16'">
          <GetCellDisplayedText />
        </div>
        <div v-if="filePageId == '3-17'">
          <GetCellValueByCellName />
        </div>
        <div v-if="filePageId == '3-18'">
          <GetIntersectionOfTwoRanges />
        </div>
        <div v-if="filePageId == '3-19'">
          <HideCellContent />
        </div>
        <div v-if="filePageId == '3-20'">
          <MergeCells />
        </div>
        <div v-if="filePageId == '3-21'">
          <SetCellFillPattern />
        </div>
        <div v-if="filePageId == '3-22'">
          <SetDBNumFormatting />
        </div>
        <div v-if="filePageId == '3-23'">
          <ShrinkTextToFitInACell />
        </div>
        <div v-if="filePageId == '3-24'">
          <TraverseCellsValue />
        </div>
        <div v-if="filePageId == '3-25'">
          <UngroupExcelCells />
        </div>
        <div v-if="filePageId == '3-26'">
          <UnmergeCells />
        </div>
        <div v-if="filePageId == '3-27'">
          <UseExplicitLineBreaks />
        </div>
        <div v-if="filePageId == '3-28'">
          <WrapOrUnwrapTextInCells />
        </div>
        <!--04_rowsColumns -->
        <div v-if="filePageId == '4-1'">
          <AutoFitColumnInRange />
        </div>
        <div v-if="filePageId == '4-2'">
          <AutoFitRowInRange />
        </div>
        <div v-if="filePageId == '4-3'">
          <CheckAutoFitRowOrColumn />
        </div>
        <div v-if="filePageId == '4-4'">
          <CheckRowOrColumnIsHidden />
        </div>
        <div v-if="filePageId == '4-5'">
          <CopyColumns />
        </div>
        <div v-if="filePageId == '4-6'">
          <CopyRows />
        </div>
        <div v-if="filePageId == '4-7'">
          <CopySingleColumnAndRow />
        </div>
        <div v-if="filePageId == '4-8'">
          <CopyWithOptions />
        </div>
        <div v-if="filePageId == '4-9'">
          <DeleteBlankRowsAndColumns />
        </div>
        <div v-if="filePageId == '4-10'">
          <DeleteMultipleRowsAndColumns />
        </div>
        <div v-if="filePageId == '4-11'">
          <GetDefaultRowAndColumnCount />
        </div>
        <div v-if="filePageId == '4-12'">
          <GroupRowsAndColumns />
        </div>
        <div v-if="filePageId == '4-13'">
          <HideOrShowRowColumnHeaders />
        </div>
        <div v-if="filePageId == '4-14'">
          <HideRowsAndColumns />
        </div>
        <div v-if="filePageId == '4-15'">
          <InsertRowsAndColumns />
        </div>
        <div v-if="filePageId == '4-16'">
          <RemoveRowBasedOnKeyword />
        </div>
        <div v-if="filePageId == '4-17'">
          <SetColumnWithInPixels />
        </div>
        <div v-if="filePageId == '4-18'">
          <SetDefaultColumnWidth />
        </div>
        <div v-if="filePageId == '4-19'">
          <SetDefaultRowAndColumnStyle />
        </div>
        <div v-if="filePageId == '4-20'">
          <SetDefaultRowHeight />
        </div>
        <div v-if="filePageId == '4-21'">
          <SetHeightAndWidth />
        </div>
        <div v-if="filePageId == '4-22'">
          <SetSummaryColumnDirection />
        </div>
        <div v-if="filePageId == '4-23'">
          <SetSummaryRowDirection />
        </div>
        <div v-if="filePageId == '4-24'">
          <UnhideRowsAndColumns />
        </div>
        <!-- 05_images -->
        <div v-if="filePageId == '5-1'">
          <AlignPictureWithinCell />
        </div>
        <div v-if="filePageId == '5-2'">
          <CompressPictures />
        </div>
        <div v-if="filePageId == '5-3'">
          <CopyPicture />
        </div>
        <div v-if="filePageId == '5-4'">
          <CroppedPositionOfPicture />
        </div>
        <div v-if="filePageId == '5-5'">
          <DeleteAllImages />
        </div>
        <div v-if="filePageId == '5-6'">
          <InsertExcelBackgroundImage />
        </div>
        <div v-if="filePageId == '5-7'">
          <LocateImages />
        </div>
        <div v-if="filePageId == '5-8'">
          <PictureOffset />
        </div>
        <div v-if="filePageId == '5-9'">
          <PictureRefRange />
        </div>
        <div v-if="filePageId == '5-10'">
          <ReadImages />
        </div>
        <div v-if="filePageId == '5-11'">
          <RemovePictureBorder />
        </div>
        <div v-if="filePageId == '5-12'">
          <ResetSizeAndPositionForImage />
        </div>
        <div v-if="filePageId == '5-13'">
          <SetImageOffsetOfChart />
        </div>
        <div v-if="filePageId == '5-14'">
          <WriteImages />
        </div>
        <!--06_Comments-->
        <div v-if="filePageId == '6-1'">
          <AddCommentWithAuthor />
        </div>
        <div v-if="filePageId == '6-2'">
          <AddCommentWithPicture />
        </div>
        <div v-if="filePageId == '6-3'">
          <EditExcelComment />
        </div>
        <div v-if="filePageId == '6-4'">
          <HideOrShowComment />
        </div>
        <div v-if="filePageId == '6-5'">
          <ReadComment />
        </div>
        <div v-if="filePageId == '6-6'">
          <RemoveComment />
        </div>
        <div v-if="filePageId == '6-7'">
          <SetCommentFillColor />
        </div>
        <div v-if="filePageId == '6-8'">
          <SetCommentTextRotation />
        </div>
        <div v-if="filePageId == '6-9'">
          <SetPositionAndAlignment />
        </div>
        <div v-if="filePageId == '6-10'">
          <WriteComment />
        </div>
        <!--07_Conversion-->
        <div v-if="filePageId == '7-1'">
          <ChartSheetToSVG />
        </div>
        <div v-if="filePageId == '7-2'">
          <CSVToExcel />
        </div>
        <div v-if="filePageId == '7-3'">
          <CSVToPDF />
        </div>
        <div v-if="filePageId == '7-4'">
          <EachWorksheetToDifferentPDF />
        </div>
        <div v-if="filePageId == '7-5'">
          <FitWidthWhenConvertToPDF />
        </div>
        <div v-if="filePageId == '7-6'">
          <HtmlToExcel />
        </div>
        <div v-if="filePageId == '7-7'">
          <LoadSaveEtAndETT />
        </div>
        <div v-if="filePageId == '7-8'">
          <OfficeOpenXMLToExcel />
        </div>
        <div v-if="filePageId == '7-9'">
          <SelectedRangeToPDF />
        </div>
        <div v-if="filePageId == '7-10'">
          <SheetToImage />
        </div>
        <div v-if="filePageId == '7-11'">
          <SpecificCellsToImage />
        </div>
        <div v-if="filePageId == '7-12'">
          <SpecifyFontDirectory />
        </div>
        <div v-if="filePageId == '7-13'">
          <ToCSV />
        </div>
        <div v-if="filePageId == '7-14'">
          <ToCSVWithFilteredValue />
        </div>
        <div v-if="filePageId == '7-15'">
          <ToHtml />
        </div>
        <div v-if="filePageId == '7-16'">
          <ToHtmlStream />
        </div>
        <div v-if="filePageId == '7-17'">
          <ToImageWithoutWhiteSpace />
        </div>
        <div v-if="filePageId == '7-18'">
          <ToODS />
        </div>
        <div v-if="filePageId == '7-19'">
          <ToOFD />
        </div>
        <div v-if="filePageId == '7-20'">
          <ToOfficeOpenXML />
        </div>
        <div v-if="filePageId == '7-21'">
          <ToPDF />
        </div>
        <div v-if="filePageId == '7-22'">
          <ToPDFA1B />
        </div>
        <div v-if="filePageId == '7-23'">
          <ToPdfSimply />
        </div>
        <div v-if="filePageId == '7-24'">
          <ToPdfWithChangePageSize />
        </div>
        <div v-if="filePageId == '7-25'">
          <ToPostScript />
        </div>
        <div v-if="filePageId == '7-26'">
          <ToSVG />
        </div>
        <div v-if="filePageId == '7-27'">
          <ToText />
        </div>
        <div v-if="filePageId == '7-28'">
          <ToXPS />
        </div>
        <div v-if="filePageId == '7-29'">
          <WorkbookToHTML />
        </div>
        <div v-if="filePageId == '7-30'">
          <XLSToXLSM />
        </div>
        <!--08_FilteringAndValidation-->
        <div v-if="filePageId == '8-1'">
          <AutofilterBlank />
        </div>
        <div v-if="filePageId == '8-2'">
          <AutofilterNonBlank />
        </div>
        <div v-if="filePageId == '8-3'">
          <CreateFilter />
        </div>
        <div v-if="filePageId == '8-4'">
          <FilterCellsByString />
        </div>
        <div v-if="filePageId == '8-5'">
          <GetSettingsOfDataValidation />
        </div>
        <div v-if="filePageId == '8-6'">
          <ListDataValidation />
        </div>
        <div v-if="filePageId == '8-7'">
          <RemoveAutoFilters />
        </div>
        <div v-if="filePageId == '8-8'">
          <RemoveDataValidation />
        </div>
        <div v-if="filePageId == '8-9'">
          <SetDataValidationOnSeparateSheet />
        </div>
        <div v-if="filePageId == '8-10'">
          <TimeDataValidation />
        </div>
        <div v-if="filePageId == '8-11'">
          <VerifyDataByValidation />
        </div>
        <div v-if="filePageId == '8-12'">
          <WholeNumberDataValidation />
        </div>
        <!-- 09_Chart -->
        <div v-if="filePageId == '9-1'">
          <AddDataTable />
        </div>
        <div v-if="filePageId == '9-2'">
          <AddErrorBars />
        </div>
        <div v-if="filePageId == '9-3'">
          <AddPictureInChart />
        </div>
        <div v-if="filePageId == '9-4'">
          <AddTextBox />
        </div>
        <div v-if="filePageId == '9-5'">
          <AddTrendline />
        </div>
        <div v-if="filePageId == '9-6'">
          <AjustBarSpace />
        </div>
        <div v-if="filePageId == '9-7'">
          <ApplySoftEdgesEffect />
        </div>
        <div v-if="filePageId == '9-8'">
          <ChangeChartSizeAndPosition />
        </div>
        <div v-if="filePageId == '9-9'">
          <ChangeDataLabel />
        </div>
        <div v-if="filePageId == '9-10'">
          <ChangeDataRange />
        </div>
        <div v-if="filePageId == '9-11'">
          <ChangeMajorGridlines />
        </div>
        <div v-if="filePageId == '9-12'">
          <ChangeSeriesColor />
        </div>
        <div v-if="filePageId == '9-13'">
          <ChartAxisTitle />
        </div>
        <div v-if="filePageId == '9-14'">
          <ChartToImage />
        </div>
        <div v-if="filePageId == '9-15'">
          <ClusteredBar />
        </div>
        <div v-if="filePageId == '9-16'">
          <ClusteredBar_3D />
        </div>
        <div v-if="filePageId == '9-17'">
          <ClusteredColumn />
        </div>
        <div v-if="filePageId == '9-18'">
          <CreateBoxAndWhiskerChart />
        </div>
        <div v-if="filePageId == '9-19'">
          <CreateBubbleChart />
        </div>
        <div v-if="filePageId == '9-20'">
          <CreateChartBasedOnPivotTable />
        </div>
        <div v-if="filePageId == '9-21'">
          <CreateChartWithoutRangeData />
        </div>
        <div v-if="filePageId == '9-22'">
          <CreateCustomChart />
        </div>
        <div v-if="filePageId == '9-23'">
          <CreateDoughnutChart />
        </div>
        <div v-if="filePageId == '9-24'">
          <CreateFunnelChart />
        </div>
        <div v-if="filePageId == '9-25'">
          <CreateHistogramChart />
        </div>
        <div v-if="filePageId == '9-26'">
          <CreateMultiLevelChart />
        </div>
        <div v-if="filePageId == '9-27'">
          <CreateParetoChart />
        </div>
        <div v-if="filePageId == '9-28'">
          <CreatePivotChart />
        </div>
        <div v-if="filePageId == '9-29'">
          <CreateRadarChart />
        </div>
        <div v-if="filePageId == '9-30'">
          <CreateSunBurstChart />
        </div>
        <div v-if="filePageId == '9-31'">
          <CreateTreeMapChart />
        </div>
        <div v-if="filePageId == '9-32'">
          <CreateWaterfallChart />
        </div>
        <div v-if="filePageId == '9-33'">
          <CustomDataMarker />
        </div>
        <div v-if="filePageId == '9-34'">
          <DataCallout />
        </div>
        <div v-if="filePageId == '9-35'">
          <DeleteLegend />
        </div>
        <div v-if="filePageId == '9-36'">
          <DiscontinuousData />
        </div>
        <div v-if="filePageId == '9-37'">
          <EditLineChart />
        </div>
        <div v-if="filePageId == '9-38'">
          <ExplodedDoughnut />
        </div>
        <div v-if="filePageId == '9-39'">
          <ExtractTrendline />
        </div>
        <div v-if="filePageId == '9-40'">
          <FillChartElementWithPicture />
        </div>
        <div v-if="filePageId == '9-41'">
          <FormatAxis />
        </div>
        <div v-if="filePageId == '9-42'">
          <GaugeChart />
        </div>
        <div v-if="filePageId == '9-43'">
          <GetCategoryLabels />
        </div>
        <div v-if="filePageId == '9-44'">
          <GetChartDataPointValues />
        </div>
        <div v-if="filePageId == '9-45'">
          <GetWorksheetOfChart />
        </div>
        <div v-if="filePageId == '9-46'">
          <HideMajorGridlines />
        </div>
        <div v-if="filePageId == '9-47'">
          <Line />
        </div>
        <div v-if="filePageId == '9-48'">
          <Pie />
        </div>
        <div v-if="filePageId == '9-49'">
          <PyramidColumn />
        </div>
        <div v-if="filePageId == '9-50'">
          <RemoveChart />
        </div>
        <div v-if="filePageId == '9-51'">
          <ResizeAndMoveChart />
        </div>
        <div v-if="filePageId == '9-52'">
          <RichTextForDataLabel />
        </div>
        <div v-if="filePageId == '9-53'">
          <Rotate3DChart />
        </div>
        <div v-if="filePageId == '9-54'">
          <ScatterChart />
        </div>
        <div v-if="filePageId == '9-55'">
          <SetAndFormatDataLabel />
        </div>
        <div v-if="filePageId == '9-56'">
          <SetBorderColorAndStyle />
        </div>
        <div v-if="filePageId == '9-57'">
          <SetBorderWidthOfMarker />
        </div>
        <div v-if="filePageId == '9-58'">
          <SetChartBackgroundColor />
        </div>
        <div v-if="filePageId == '9-59'">
          <SetColorForChartArea />
        </div>
        <div v-if="filePageId == '9-60'">
          <SetFont />
        </div>
        <div v-if="filePageId == '9-61'">
          <SetFontForLegendAndDataTable />
        </div>
        <div v-if="filePageId == '9-62'">
          <SetFontForTitleAndAxis />
        </div>
        <div v-if="filePageId == '9-63'">
          <SetLegendBackgroundColor />
        </div>
        <div v-if="filePageId == '9-64'">
          <SetNumberFormatOfTrendline />
        </div>
        <div v-if="filePageId == '9-65'">
          <ShowLeaderLine />
        </div>
        <div v-if="filePageId == '9-66'">
          <SparkLine />
        </div>
        <div v-if="filePageId == '9-67'">
          <StackedColumn />
        </div>
        <!-- 10_Shape -->
        <div v-if="filePageId == '10-1'">
          <AddArrowLineToExcelFile />
        </div>
        <div v-if="filePageId == '10-2'">
          <AddLineShape />
        </div>
        <div v-if="filePageId == '10-3'">
          <AddOvalShape />
        </div>
        <div v-if="filePageId == '10-4'">
          <AddRectangleShape />
        </div>
        <div v-if="filePageId == '10-5'">
          <AddSpinnerControl />
        </div>
        <div v-if="filePageId == '10-6'">
          <AdjustArrowPolylinePosition />
        </div>
        <div v-if="filePageId == '10-7'">
          <CopyShapes />
        </div>
        <div v-if="filePageId == '10-8'">
          <DeleteAllShapes />
        </div>
        <div v-if="filePageId == '10-9'">
          <DeleteParticularShape />
        </div>
        <div v-if="filePageId == '10-10'">
          <DrawOneLineThroughTwoPoints />
        </div>
        <div v-if="filePageId == '10-11'">
          <ExtractTextImageFromShape />
        </div>
        <div v-if="filePageId == '10-12'">
          <GetShapeLinkedCellRange />
        </div>
        <div v-if="filePageId == '10-13'">
          <HideOrUnhideShape />
        </div>
        <div v-if="filePageId == '10-14'">
          <InsertShapesToExcelSheet />
        </div>
        <div v-if="filePageId == '10-15'">
          <ModifyShadowStyleForShape />
        </div>
        <div v-if="filePageId == '10-16'">
          <SetShadowStyleForShape />
        </div>
        <div v-if="filePageId == '10-17'">
          <SetShapeOrder />
        </div>
        <div v-if="filePageId == '10-18'">
          <ShapeToImage />
        </div>
        <div v-if="filePageId == '10-19'">
          <TillPicAsTextureInShape />
        </div>
        <!--11_Formatting-->
        <div v-if="filePageId == '11-1'">
          <ApplyBuiltInStyles />
        </div>
        <div v-if="filePageId == '11-2'">
          <ApplyColorScalesToDataRange />
        </div>
        <div v-if="filePageId == '11-3'">
          <ApplyConditionalFormatting />
        </div>
        <div v-if="filePageId == '11-4'">
          <ApplyDataBarsToCellRange />
        </div>
        <div v-if="filePageId == '11-5'">
          <ApplyGradientFillEffects />
        </div>
        <div v-if="filePageId == '11-6'">
          <ApplyIconSetsToCellRange />
        </div>
        <div v-if="filePageId == '11-7'">
          <ColorsAndPalette />
        </div>
        <div v-if="filePageId == '11-8'">
          <ConditionalFormatRuntime />
        </div>
        <div v-if="filePageId == '11-9'">
          <ConditionallyFormatDate />
        </div>
        <div v-if="filePageId == '11-10'">
          <CreateFormulaConditionalFormat />
        </div>
        <div v-if="filePageId == '11-11'">
          <FontStyles />
        </div>
        <div v-if="filePageId == '11-12'">
          <ForegroundAndBackground />
        </div>
        <div v-if="filePageId == '11-13'">
          <FormatAColumn />
        </div>
        <div v-if="filePageId == '11-14'">
          <FormatARow />
        </div>
        <div v-if="filePageId == '11-15'">
          <FormatCellsWithStyle />
        </div>
        <div v-if="filePageId == '11-16'">
          <GetColorArgbData />
        </div>
        <div v-if="filePageId == '11-17'">
          <GetStyleSetStyle />
        </div>
        <div v-if="filePageId == '11-18'">
          <HighlightAverageValues />
        </div>
        <div v-if="filePageId == '11-19'">
          <HighlightDuplicateUniqueValues />
        </div>
        <div v-if="filePageId == '11-20'">
          <HighlightRankedValues />
        </div>
        <div v-if="filePageId == '11-21'">
          <Indentation />
        </div>
        <div v-if="filePageId == '11-22'">
          <Interior />
        </div>
        <div v-if="filePageId == '11-23'">
          <MakeCellActive />
        </div>
        <div v-if="filePageId == '11-24'">
          <NumberStyles />
        </div>
        <div v-if="filePageId == '11-25'">
          <SetBorder />
        </div>
        <div v-if="filePageId == '11-26'">
          <SetBorderToDataBar />
        </div>
        <div v-if="filePageId == '11-27'">
          <SetConditionalFormatFormula />
        </div>
        <div v-if="filePageId == '11-28'">
          <SetRowColorByConditionalFormat />
        </div>
        <div v-if="filePageId == '11-29'">
          <SetTrafficLightsIcons />
        </div>
        <div v-if="filePageId == '11-30'">
          <SimpleConditionalFormatting />
        </div>
        <div v-if="filePageId == '11-31'">
          <TextAlign />
        </div>
        <div v-if="filePageId == '11-32'">
          <TextDirection />
        </div>
        <div v-if="filePageId == '11-33'">
          <UsePredefinedStyles />
        </div>
        <div v-if="filePageId == '11-34'">
          <UsingStyleObject />
        </div>
        <div v-if="filePageId == '11-35'">
          <VariousConditionalFormatting />
        </div>
        <!--12_Formulas-->
        <div v-if="filePageId == '12-1'">
          <InsertFormulaWithNamedRange />
        </div>
        <div v-if="filePageId == '12-2'">
          <ReadFormulas />
        </div>
        <div v-if="filePageId == '12-3'">
          <RegisterAddInFunction />
        </div>
        <div v-if="filePageId == '12-4'">
          <RemoveFormulasButKeepValues />
        </div>
        <div v-if="filePageId == '12-5'">
          <SubTotalFormula />
        </div>
        <div v-if="filePageId == '12-6'">
          <UseArrayFormulas />
        </div>
        <div v-if="filePageId == '12-7'">
          <UseArrayR1C1Formula />
        </div>
        <div v-if="filePageId == '12-8'">
          <UseR1C1Formula />
        </div>
        <div v-if="filePageId == '12-9'">
          <WriteFormulas />
        </div>
        <!--13_HeaderFooter-->
        <div v-if="filePageId == '13-1'">
          <ChangeFontAndSize />
        </div>
        <div v-if="filePageId == '13-2'">
          <DifferentHeaderFooter />
        </div>
        <div v-if="filePageId == '13-3'">
          <DifferentHeaderFooterOnFirstPage />
        </div>
        <div v-if="filePageId == '13-4'">
          <ImageHeaderFooter />
        </div>
        <div v-if="filePageId == '13-5'">
          <SetHeaderFooter />
        </div>
        <!--14_Hyperlinks-->
        <div v-if="filePageId == '14-1'">
          <AddHyperlinkToText />
        </div>
        <div v-if="filePageId == '14-2'">
          <AddImageHyperlink />
        </div>
        <div v-if="filePageId == '14-3'">
          <GetHyperLinkType />
        </div>
        <div v-if="filePageId == '14-4'">
          <LinkToOtherSheetCell />
        </div>
        <div v-if="filePageId == '14-5'">
          <ModifyHyperlink />
        </div>
        <div v-if="filePageId == '14-6'">
          <ReadHyperlinks />
        </div>
        <div v-if="filePageId == '14-7'">
          <RemoveHyperlinks />
        </div>
        <div v-if="filePageId == '14-8'">
          <RetrieveExternalFileHyperlinks />
        </div>
        <div v-if="filePageId == '14-9'">
          <WriteHyperlinks />
        </div>
        <div v-if="filePageId == '14-10'">
          <LinkToExternalFile />
        </div>
        <!--15_MarkerDesigner-->
        <div v-if="filePageId == '15-1'">
          <AddVariableArray />
        </div>
        <!--16_Namedranges-->
        <div v-if="filePageId == '16-1'">
          <FormatNamedRangeCells />
        </div>
        <div v-if="filePageId == '16-2'">
          <GetAllNamedRange />
        </div>
        <div v-if="filePageId == '16-3'">
          <GetNamedRangeAddress />
        </div>
        <div v-if="filePageId == '16-4'">
          <GetSpecificNamedRange />
        </div>
        <div v-if="filePageId == '16-5'">
          <MergeNamedRangeCells />
        </div>
        <div v-if="filePageId == '16-6'">
          <NamedRanges />
        </div>
        <div v-if="filePageId == '16-7'">
          <RemoveNamedRange />
        </div>
        <div v-if="filePageId == '16-8'">
          <RenameNamedRange />
        </div>
        <div v-if="filePageId == '16-9'">
          <ScopedNamedRange />
        </div>
        <div v-if="filePageId == '16-10'">
          <SetFormulaWithNamedRange />
        </div>
        <!--17_Oleobjects-->
        <div v-if="filePageId == '17-1'">
          <ExtractOLEObjects />
        </div>
        <div v-if="filePageId == '17-2'">
          <InsertOLEObjects />
        </div>
        <div v-if="filePageId == '17-3'">
          <InsertWavFileOLEObject />
        </div>
        <!--18_Pagesetup-->
        <div v-if="filePageId == '18-1'">
          <GetExcelPaperDimensions />
        </div>
        <div v-if="filePageId == '18-2'">
          <GetExcelVersion />
        </div>
        <div v-if="filePageId == '18-3'">
          <SetExcelPageOrderType />
        </div>
        <div v-if="filePageId == '18-4'">
          <SetExcelPaperSize />
        </div>
        <div v-if="filePageId == '18-5'">
          <SetFirstPageNumber />
        </div>
        <div v-if="filePageId == '18-6'">
          <SetHeaderAndFooterMargins />
        </div>
        <div v-if="filePageId == '18-7'">
          <SetMarginsOfExcelSheet />
        </div>
        <div v-if="filePageId == '18-8'">
          <SetOtherPrintingOptions />
        </div>
        <div v-if="filePageId == '18-9'">
          <SetPageOrientation />
        </div>
        <div v-if="filePageId == '18-10'">
          <SetPrintAreaOfXlsFile />
        </div>
        <div v-if="filePageId == '18-11'">
          <SetPrintQualityOfXlsFile />
        </div>
        <div v-if="filePageId == '18-12'">
          <SetPrintTitleOfXlsFile />
        </div>
        <div v-if="filePageId == '18-13'">
          <SetSheetFitToPageProperty />
        </div>
        <div v-if="filePageId == '18-14'">
          <SetXlsSheetCenterOnPage />
        </div>
        <!--19_Pivottables-->
        <div v-if="filePageId == '19-1'">
          <ChangeDataSource />
        </div>
        <div v-if="filePageId == '19-2'">
          <ClearPivotFields />
        </div>
        <div v-if="filePageId == '19-3'">
          <ConsolidationFunctions />
        </div>
        <div v-if="filePageId == '19-4'">
          <CreatePivotTable />
        </div>
        <div v-if="filePageId == '19-5'">
          <DisablePivotTableRibbon />
        </div>
        <div v-if="filePageId == '19-6'">
          <ExpandOrCollapseRows />
        </div>
        <div v-if="filePageId == '19-7'">
          <FormatDataField />
        </div>
        <div v-if="filePageId == '19-8'">
          <FormattingAppearance />
        </div>
        <div v-if="filePageId == '19-9'">
          <GetPivotTableRefreshedInfo />
        </div>
        <div v-if="filePageId == '19-10'">
          <RefreshPivotTable />
        </div>
        <div v-if="filePageId == '19-11'">
          <RepeatItemLabels />
        </div>
        <div v-if="filePageId == '19-12'">
          <SetFormatOptions />
        </div>
        <div v-if="filePageId == '19-13'">
          <SetPivotFieldFormat />
        </div>
        <div v-if="filePageId == '19-14'">
          <ShowDataFieldInRow />
        </div>
        <div v-if="filePageId == '19-15'">
          <ShowSubTotals />
        </div>
        <div v-if="filePageId == '19-16'">
          <SortPivotTable />
        </div>
        <div v-if="filePageId == '19-17'">
          <UpdateDataSource />
        </div>
        <!-- 20_Security -->
        <div v-if="filePageId == '20-1'">
          <DetectProtection />
        </div>
        <div v-if="filePageId == '20-2'">
          <HideFormulas />
        </div>
        <div v-if="filePageId == '20-3'">
          <LockSpecificCellInNewExcel />
        </div>
        <div v-if="filePageId == '20-4'">
          <LockSpecificColumnInNewExcel />
        </div>
        <div v-if="filePageId == '20-5'">
          <LockSpecificRowInNewExcel />
        </div>
        <div v-if="filePageId == '20-6'">
          <ProtectCell />
        </div>
        <div v-if="filePageId == '20-7'">
          <ProtectWithEditableRange />
        </div>
        <div v-if="filePageId == '20-8'">
          <ProtectWorkbook />
        </div>
        <div v-if="filePageId == '20-9'">
          <RemoveAllDigitalSignatures />
        </div>
        <div v-if="filePageId == '20-10'">
          <UnlockProtectSheet />
        </div>
        <div v-if="filePageId == '20-11'">
          <UnlockSimpleSheet />
        </div>

        <!-- 21-Texboxes -->
        <div v-if="filePageId == '21-1'">
          <ExtractTextFromATextbox />
        </div>
        <div v-if="filePageId == '21-2'">
          <GetTextBoxByName />
        </div>
        <div v-if="filePageId == '21-3'">
          <ManipulateTextBox />
        </div>
        <div v-if="filePageId == '21-4'">
          <RemoveBorderlineOfTextbox />
        </div>
        <div v-if="filePageId == '21-5'">
          <ReplaceTextInTextBox />
        </div>
        <div v-if="filePageId == '21-6'">
          <SetFontAndBackground />
        </div>
        <div v-if="filePageId == '21-7'">
          <SetInternalMarginOfTextbox />
        </div>
        <div v-if="filePageId == '21-8'">
          <TextBoxWithWrapText />
        </div>
        <!-- 22_worksheets -->
        <div v-if="filePageId == '22-1'">
          <ActivateWorksheet />
        </div>
        <div v-if="filePageId == '22-2'">
          <AddPageBreakInXlsFile />
        </div>
        <div v-if="filePageId == '22-3'">
          <AddWorksheet />
        </div>
        <div v-if="filePageId == '22-4'">
          <ApplyStyleToWorksheet />
        </div>
        <div v-if="filePageId == '22-5'">
          <CheckDialogSheetInXlsFile />
        </div>
        <div v-if="filePageId == '22-6'">
          <CopySheetToAnotherXlsFile />
        </div>
        <div v-if="filePageId == '22-7'">
          <CopySheetWithinWorkbook />
        </div>
        <div v-if="filePageId == '22-8'">
          <CopyVisibleSheets />
        </div>
        <div v-if="filePageId == '22-9'">
          <CopyWorksheet />
        </div>
        <div v-if="filePageId == '22-10'">
          <DetectEmptyWorksheet />
        </div>
        <div v-if="filePageId == '22-11'">
          <FillDataInWorksheet />
        </div>
        <div v-if="filePageId == '22-12'">
          <FreezePanes />
        </div>
        <div v-if="filePageId == '22-13'">
          <GetFreezePaneRange />
        </div>
        <div v-if="filePageId == '22-14'">
          <GetListOfFontsUsed />
        </div>
        <div v-if="filePageId == '22-15'">
          <GetPaperSize />
        </div>
        <div v-if="filePageId == '22-16'">
          <GetWorksheetNames />
        </div>
        <div v-if="filePageId == '22-17'">
          <HideOrShowWorksheet />
        </div>
        <div v-if="filePageId == '22-18'">
          <HideTab />
        </div>
        <div v-if="filePageId == '22-19'">
          <HideZeroValues />
        </div>
        <div v-if="filePageId == '22-20'">
          <LinkToContentProperty />
        </div>
        <div v-if="filePageId == '22-21'">
          <MoveWorksheet />
        </div>
        <div v-if="filePageId == '22-22'">
          <PageBreakPreview />
        </div>
        <div v-if="filePageId == '22-23'">
          <RemovePageBreak />
        </div>
        <div v-if="filePageId == '22-24'">
          <RemoveWorksheet />
        </div>
        <div v-if="filePageId == '22-25'">
          <SetPageBreak />
        </div>
        <div v-if="filePageId == '22-26'">
          <SetTabColor />
        </div>
        <div v-if="filePageId == '22-27'">
          <SetViewMode />
        </div>
        <div v-if="filePageId == '22-28'">
          <ShowOrHideGridLine />
        </div>
        <div v-if="filePageId == '22-29'">
          <ShowTab />
        </div>
        <div v-if="filePageId == '22-30'">
          <SplitWorksheetIntoPanes />
        </div>
        <div v-if="filePageId == '22-31'">
          <UnfreezeExcelPanes />
        </div>
        <div v-if="filePageId == '22-32'">
          <VerifyProtectedWorksheet />
        </div>
        <div v-if="filePageId == '22-33'">
          <ZoomFactor />
        </div>

        <!-- 23_workbook -->
        <div v-if="filePageId == '23-1'">
          <AccessDocumentProperties />
        </div>
        <div v-if="filePageId == '23-2'">
          <AddCustomProperties />
        </div>
        <div v-if="filePageId == '23-3'">
          <DecryptWorkbook />
        </div>
        <div v-if="filePageId == '23-4'">
          <DetectExcelVersion />
        </div>
        <div v-if="filePageId == '23-5'">
          <DetectVBAMacros />
        </div>
        <div v-if="filePageId == '23-6'">
          <EncryptWorkbook />
        </div>
        <div v-if="filePageId == '23-7'">
          <GetProperties />
        </div>
        <div v-if="filePageId == '23-8'">
          <HideWindow />
        </div>
        <div v-if="filePageId == '23-9'">
          <LoadAndSaveFileWithMacro />
        </div>
        <div v-if="filePageId == '23-10'">
          <MergeExcelFiles />
        </div>
        <div v-if="filePageId == '23-11'">
          <OpenEncryptedFile />
        </div>
        <div v-if="filePageId == '23-12'">
          <OpenFiles />
        </div>
        <div v-if="filePageId == '23-13'">
          <ReadStream />
        </div>
        <div v-if="filePageId == '23-14'">
          <RemoveCustomProperties />
        </div>
        <div v-if="filePageId == '23-15'">
          <SaveFiles />
        </div>
        <div v-if="filePageId == '23-16'">
          <SaveStream />
        </div>
        <div v-if="filePageId == '23-17'">
          <SetExcelCalculationMode />
        </div>
        <div v-if="filePageId == '23-18'">
          <SetMargins />
        </div>
        <div v-if="filePageId == '23-19'">
          <SetTheme />
        </div>
        <!--24_TrackChanges-->
        <div v-if="filePageId == '24-1'">
          <AcceptOrRejectTrackedChanges />
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref } from "vue";
import SpireTree from "@/components/SpireTree.vue";
import Description from "@/components/Description.vue";

//01_quickguide
import CreateAnExcelWithFiveSheets from "@/components/01_quickguide/CreateAnExcelWithFiveSheets.vue";
import CreateAnExcelWithOneSheet from "@/components/01_quickguide/CreateAnExcelWithOneSheet.vue";
import CreateFiftyExcelFiles from "@/components/01_quickguide/CreateFiftyExcelFiles.vue";
import HelloWorld from "@/components/01_quickguide/HelloWorld.vue";
import OpenExistingFile from "@/components/01_quickguide/OpenExistingFile.vue";

//02_data
import AddLabelControl from "@/components/02_data/AddLabelControl.vue";
import AddListBoxControl from "@/components/02_data/AddListBoxControl.vue";
import AddScrollBarControl from "@/components/02_data/AddScrollBarControl.vue";
import AddTableWithFilter from "@/components/02_data/AddTableWithFilter.vue";
import AddTotalRowToTable from "@/components/02_data/AddTotalRowToTable.vue";
import ApplySubscriptAndSuperscript from "@/components/02_data/ApplySubscriptAndSuperscript.vue";
import CloneExcelFontStyle from "@/components/02_data/CloneExcelFontStyle.vue";
import CopyCellsRange from "@/components/02_data/CopyCellsRange.vue";
import CopyDataWithStyle from "@/components/02_data/CopyDataWithStyle.vue";
import CopyOnlyFormulaValue from "@/components/02_data/CopyOnlyFormulaValue.vue";
import CreateNestedGroup from "@/components/02_data/CreateNestedGroup.vue";
import CreateTable from "@/components/02_data/CreateTable.vue";
import DataSorting from "@/components/02_data/DataSorting.vue";
import DataValidation from "@/components/02_data/DataValidation.vue";
import ExpandAndCollapseGroups from "@/components/02_data/ExpandAndCollapseGroups.vue";
import FindAndReplaceData from "@/components/02_data/FindAndReplaceData.vue";
import FindDataInSpecificRange from "@/components/02_data/FindDataInSpecificRange.vue";
import FindStringAndNumber from "@/components/02_data/FindStringAndNumber.vue";
import FormatTable from "@/components/02_data/FormatTable.vue";
import InsertControls from "@/components/02_data/InsertControls.vue";
import InsertHtmlStringIntoCell from "@/components/02_data/InsertHtmlStringIntoCell.vue";
import ReplaceAndHighlight from "@/components/02_data/ReplaceAndHighlight.vue";
import RetrieveAndExtractData from "@/components/02_data/RetrieveAndExtractData.vue";
import SplitDataIntoMultipleColumns from "@/components/02_data/SplitDataIntoMultipleColumns.vue";
import Subtotal from "@/components/02_data/Subtotal.vue";
import WriteRichText from "@/components/02_data/WriteRichText.vue";

//03_cells
import AccessCell from "@/components/03_cells/AccessCell.vue";
import ApplyMultipleFontsInSingleCell from "@/components/03_cells/ApplyMultipleFontsInSingleCell.vue";
import AutoFitBasedOnCellValue from "@/components/03_cells/AutoFitBasedOnCellValue.vue";
import ConvertTextToNubmer from "@/components/03_cells/ConvertTextToNubmer.vue";
import CopyCellFormat from "@/components/03_cells/CopyCellFormat.vue";
import CountNumberOfCells from "@/components/03_cells/CountNumberOfCells.vue";
import CutCellsToOtherPosition from "@/components/03_cells/CutCellsToOtherPosition.vue";
import DetectMergedCells from "@/components/03_cells/DetectMergedCells.vue";
import DuplicateCellRange from "@/components/03_cells/DuplicateCellRange.vue";
import EmptyCell from "@/components/03_cells/EmptyCell.vue";
import FilterCellsByCellColor from "@/components/03_cells/FilterCellsByCellColor.vue";
import FindCellsWithStyleName from "@/components/03_cells/FindCellsWithStyleName.vue";
import FindFormulaCells from "@/components/03_cells/FindFormulaCells.vue";
import GetCellAddress from "@/components/03_cells/GetCellAddress.vue";
import GetCellDataType from "@/components/03_cells/GetCellDataType.vue";
import GetCellDisplayedText from "@/components/03_cells/GetCellDisplayedText.vue";
import GetCellValueByCellName from "@/components/03_cells/GetCellValueByCellName.vue";
import GetIntersectionOfTwoRanges from "@/components/03_cells/GetIntersectionOfTwoRanges.vue";
import HideCellContent from "@/components/03_cells/HideCellContent.vue";
import MergeCells from "@/components/03_cells/MergeCells.vue";
import SetCellFillPattern from "@/components/03_cells/SetCellFillPattern.vue";
import SetDBNumFormatting from "@/components/03_cells/SetDBNumFormatting.vue";
import ShrinkTextToFitInACell from "@/components/03_cells/ShrinkTextToFitInACell.vue";
import TraverseCellsValue from "@/components/03_cells/TraverseCellsValue.vue";
import UngroupExcelCells from "@/components/03_cells/UngroupExcelCells.vue";
import UnmergeCells from "@/components/03_cells/UnmergeCells.vue";
import UseExplicitLineBreaks from "@/components/03_cells/UseExplicitLineBreaks.vue";
import WrapOrUnwrapTextInCells from "@/components/03_cells/WrapOrUnwrapTextInCells.vue";

//04_rowsColumns
import AutoFitColumnInRange from "@/components/04_rowsColumns/AutoFitColumnInRange.vue";
import AutoFitRowInRange from "@/components/04_rowsColumns/AutoFitRowInRange.vue";
import CheckAutoFitRowOrColumn from "@/components/04_rowsColumns/CheckAutoFitRowOrColumn.vue";
import CheckRowOrColumnIsHidden from "@/components/04_rowsColumns/CheckRowOrColumnIsHidden.vue";
import CopyColumns from "@/components/04_rowsColumns/CopyColumns.vue";
import CopyRows from "@/components/04_rowsColumns/CopyRows.vue";
import CopySingleColumnAndRow from "@/components/04_rowsColumns/CopySingleColumnAndRow.vue";
import CopyWithOptions from "@/components/04_rowsColumns/CopyWithOptions.vue";
import DeleteBlankRowsAndColumns from "@/components/04_rowsColumns/DeleteBlankRowsAndColumns.vue";
import DeleteMultipleRowsAndColumns from "@/components/04_rowsColumns/DeleteMultipleRowsAndColumns.vue";
import GetDefaultRowAndColumnCount from "@/components/04_rowsColumns/GetDefaultRowAndColumnCount.vue";
import GroupRowsAndColumns from "@/components/04_rowsColumns/GroupRowsAndColumns.vue";
import HideOrShowRowColumnHeaders from "@/components/04_rowsColumns/HideOrShowRowColumnHeaders.vue";
import HideRowsAndColumns from "@/components/04_rowsColumns/HideRowsAndColumns.vue";
import InsertRowsAndColumns from "@/components/04_rowsColumns/InsertRowsAndColumns.vue";
import RemoveRowBasedOnKeyword from "@/components/04_rowsColumns/RemoveRowBasedOnKeyword.vue";
import SetColumnWithInPixels from "@/components/04_rowsColumns/SetColumnWithInPixels.vue";
import SetDefaultColumnWidth from "@/components/04_rowsColumns/SetDefaultColumnWidth.vue";
import SetDefaultRowAndColumnStyle from "@/components/04_rowsColumns/SetDefaultRowAndColumnStyle.vue";
import SetDefaultRowHeight from "@/components/04_rowsColumns/SetDefaultRowHeight.vue";
import SetHeightAndWidth from "@/components/04_rowsColumns/SetHeightAndWidth.vue";
import SetSummaryColumnDirection from "@/components/04_rowsColumns/SetSummaryColumnDirection.vue";
import SetSummaryRowDirection from "@/components/04_rowsColumns/SetSummaryRowDirection.vue";
import UnhideRowsAndColumns from "@/components/04_rowsColumns/UnhideRowsAndColumns.vue";

//05_images
import AlignPictureWithinCell from "@/components/05_images/AlignPictureWithinCell.vue";
import CompressPictures from "@/components/05_images/CompressPictures.vue";
import CopyPicture from "@/components/05_images/CopyPicture.vue";
import CroppedPositionOfPicture from "@/components/05_images/CroppedPositionOfPicture.vue";
import DeleteAllImages from "@/components/05_images/DeleteAllImages.vue";
import InsertExcelBackgroundImage from "@/components/05_images/InsertExcelBackgroundImage.vue";
import LocateImages from "@/components/05_images/LocateImages.vue";
import PictureOffset from "@/components/05_images/PictureOffset.vue";
import PictureRefRange from "@/components/05_images/PictureRefRange.vue";
import ReadImages from "@/components/05_images/ReadImages.vue";
import RemovePictureBorder from "@/components/05_images/RemovePictureBorder.vue";
import ResetSizeAndPositionForImage from "@/components/05_images/ResetSizeAndPositionForImage.vue";
import SetImageOffsetOfChart from "@/components/05_images/SetImageOffsetOfChart.vue";
import WriteImages from "@/components/05_images/WriteImages.vue";

//06_comments
import AddCommentWithAuthor from "@/components/06_comments/AddCommentWithAuthor.vue";
import AddCommentWithPicture from "@/components/06_comments/AddCommentWithPicture.vue";
import EditExcelComment from "@/components/06_comments/EditExcelComment.vue";
import HideOrShowComment from "@/components/06_comments/HideOrShowComment.vue";
import ReadComment from "@/components/06_comments/ReadComment.vue";
import RemoveComment from "@/components/06_comments/RemoveComment.vue";
import SetCommentFillColor from "@/components/06_comments/SetCommentFillColor.vue";
import SetCommentTextRotation from "@/components/06_comments/SetCommentTextRotation.vue";
import SetPositionAndAlignment from "@/components/06_comments/SetPositionAndAlignment.vue";
import WriteComment from "@/components/06_comments/WriteComment.vue";

//07_conversion
import ChartSheetToSVG from "@/components/07_conversion/ChartSheetToSVG.vue";
import CSVToExcel from "@/components/07_conversion/CSVToExcel.vue";
import CSVToPDF from "@/components/07_conversion/CSVToPDF.vue";
import EachWorksheetToDifferentPDF from "@/components/07_conversion/EachWorksheetToDifferentPDF.vue";
import FitWidthWhenConvertToPDF from "@/components/07_conversion/FitWidthWhenConvertToPDF.vue";
import HtmlToExcel from "@/components/07_conversion/HtmlToExcel.vue";
import LoadSaveEtAndETT from "@/components/07_conversion/LoadSaveEtAndETT.vue";
import OfficeOpenXMLToExcel from "@/components/07_conversion/OfficeOpenXMLToExcel.vue";
import SelectedRangeToPDF from "@/components/07_conversion/SelectedRangeToPDF.vue";
import SheetToImage from "@/components/07_conversion/SheetToImage.vue";
import SpecificCellsToImage from "@/components/07_conversion/SpecificCellsToImage.vue";
import SpecifyFontDirectory from "@/components/07_conversion/SpecifyFontDirectory.vue";
import ToCSV from "@/components/07_conversion/ToCSV.vue";
import ToCSVWithFilteredValue from "@/components/07_conversion/ToCSVWithFilteredValue.vue";
import ToHtml from "@/components/07_conversion/ToHtml.vue";
import ToHtmlStream from "@/components/07_conversion/ToHtmlStream.vue";
import ToImageWithoutWhiteSpace from "@/components/07_conversion/ToImageWithoutWhiteSpace.vue";
import ToODS from "@/components/07_conversion/ToODS.vue";
import ToOFD from "@/components/07_conversion/ToOFD.vue";
import ToOfficeOpenXML from "@/components/07_conversion/ToOfficeOpenXML.vue";
import ToPDF from "@/components/07_conversion/ToPDF.vue";
import ToPDFA1B from "@/components/07_conversion/ToPDFA1B.vue";
import ToPdfSimply from "@/components/07_conversion/ToPdfSimply.vue";
import ToPdfWithChangePageSize from "@/components/07_conversion/ToPdfWithChangePageSize.vue";
import ToPostScript from "@/components/07_conversion/ToPostScript.vue";
import ToSVG from "@/components/07_conversion/ToSVG.vue";
import ToText from "@/components/07_conversion/ToText.vue";
import ToXPS from "@/components/07_conversion/ToXPS.vue";
import WorkbookToHTML from "@/components/07_conversion/WorkbookToHTML.vue";
import XLSToXLSM from "@/components/07_conversion/XLSToXLSM.vue";

//08_filteringAndValidation
import AutofilterBlank from "@/components/08_filteringAndValidation/AutofilterBlank.vue";
import AutofilterNonBlank from "@/components/08_filteringAndValidation/AutofilterNonBlank.vue";
import CreateFilter from "@/components/08_filteringAndValidation/CreateFilter.vue";
import FilterCellsByString from "@/components/08_filteringAndValidation/FilterCellsByString.vue";
import GetSettingsOfDataValidation from "@/components/08_filteringAndValidation/GetSettingsOfDataValidation.vue";
import ListDataValidation from "@/components/08_filteringAndValidation/ListDataValidation.vue";
import RemoveAutoFilters from "@/components/08_filteringAndValidation/RemoveAutoFilters.vue";
import RemoveDataValidation from "@/components/08_filteringAndValidation/RemoveDataValidation.vue";
import SetDataValidationOnSeparateSheet from "@/components/08_filteringAndValidation/SetDataValidationOnSeparateSheet.vue";
import TimeDataValidation from "@/components/08_filteringAndValidation/TimeDataValidation.vue";
import VerifyDataByValidation from "@/components/08_filteringAndValidation/VerifyDataByValidation.vue";
import WholeNumberDataValidation from "@/components/08_filteringAndValidation/WholeNumberDataValidation.vue";

// 09_chart
import AddDataTable from "@/components/09_chart/AddDataTable.vue";
import AddErrorBars from "@/components/09_chart/AddErrorBars.vue";
import AddPictureInChart from "@/components/09_chart/AddPictureInChart.vue";
import AddTextBox from "@/components/09_chart/AddTextBox.vue";
import AddTrendline from "@/components/09_chart/AddTrendline.vue";
import AjustBarSpace from "@/components/09_chart/AjustBarSpace.vue";
import ApplySoftEdgesEffect from "@/components/09_chart/ApplySoftEdgesEffect.vue";
import ChangeChartSizeAndPosition from "@/components/09_chart/ChangeChartSizeAndPosition.vue";
import ChangeDataLabel from "@/components/09_chart/ChangeDataLabel.vue";
import ChangeDataRange from "@/components/09_chart/ChangeDataRange.vue";
import ChangeMajorGridlines from "@/components/09_chart/ChangeMajorGridlines.vue";
import ChangeSeriesColor from "@/components/09_chart/ChangeSeriesColor.vue";
import ChartAxisTitle from "@/components/09_chart/ChartAxisTitle.vue";
import ChartToImage from "@/components/09_chart/ChartToImage.vue";
import ClusteredBar from "@/components/09_chart/ClusteredBar.vue";
import ClusteredBar_3D from "@/components/09_chart/ClusteredBar_3D.vue";
import ClusteredColumn from "@/components/09_chart/ClusteredColumn.vue";
import CreateBoxAndWhiskerChart from "@/components/09_chart/CreateBoxAndWhiskerChart.vue";
import CreateBubbleChart from "@/components/09_chart/CreateBubbleChart.vue";
import CreateChartBasedOnPivotTable from "@/components/09_chart/CreateChartBasedOnPivotTable.vue";
import CreateChartWithoutRangeData from "@/components/09_chart/CreateChartWithoutRangeData.vue";
import CreateCustomChart from "@/components/09_chart/CreateCustomChart.vue";
import CreateDoughnutChart from "@/components/09_chart/CreateDoughnutChart.vue";
import CreateFunnelChart from "@/components/09_chart/CreateFunnelChart.vue";
import CreateHistogramChart from "@/components/09_chart/CreateHistogramChart.vue";
import CreateMultiLevelChart from "@/components/09_chart/CreateMultiLevelChart.vue";
import CreateParetoChart from "@/components/09_chart/CreateParetoChart.vue";
import CreatePivotChart from "@/components/09_chart/CreatePivotChart.vue";
import CreateRadarChart from "@/components/09_chart/CreateRadarChart.vue";
import CreateSunBurstChart from "@/components/09_chart/CreateSunBurstChart.vue";
import CreateTreeMapChart from "@/components/09_chart/CreateTreeMapChart.vue";
import CreateWaterfallChart from "@/components/09_chart/CreateWaterfallChart.vue";
import CustomDataMarker from "@/components/09_chart/CustomDataMarker.vue";
import DataCallout from "@/components/09_chart/DataCallout.vue";
import DeleteLegend from "@/components/09_chart/DeleteLegend.vue";
import DiscontinuousData from "@/components/09_chart/DiscontinuousData.vue";
import EditLineChart from "@/components/09_chart/EditLineChart.vue";
import ExplodedDoughnut from "@/components/09_chart/ExplodedDoughnut.vue";
import ExtractTrendline from "@/components/09_chart/ExtractTrendline.vue";
import FillChartElementWithPicture from "@/components/09_chart/FillChartElementWithPicture.vue";
import FormatAxis from "@/components/09_chart/FormatAxis.vue";
import GaugeChart from "@/components/09_chart/GaugeChart.vue";
import GetCategoryLabels from "@/components/09_chart/GetCategoryLabels.vue";
import GetChartDataPointValues from "@/components/09_chart/GetChartDataPointValues.vue";
import GetWorksheetOfChart from "@/components/09_chart/GetWorksheetOfChart.vue";
import HideMajorGridlines from "@/components/09_chart/HideMajorGridlines.vue";
import Line from "@/components/09_chart/Line.vue";
import Pie from "@/components/09_chart/Pie.vue";
import PyramidColumn from "@/components/09_chart/PyramidColumn.vue";
import RemoveChart from "@/components/09_chart/RemoveChart.vue";
import ResizeAndMoveChart from "@/components/09_chart/ResizeAndMoveChart.vue";
import RichTextForDataLabel from "@/components/09_chart/RichTextForDataLabel.vue";
import Rotate3DChart from "@/components/09_chart/Rotate3DChart.vue";
import ScatterChart from "@/components/09_chart/ScatterChart.vue";
import SetAndFormatDataLabel from "@/components/09_chart/SetAndFormatDataLabel.vue";
import SetBorderColorAndStyle from "@/components/09_chart/SetBorderColorAndStyle.vue";
import SetBorderWidthOfMarker from "@/components/09_chart/SetBorderWidthOfMarker.vue";
import SetChartBackgroundColor from "@/components/09_chart/SetChartBackgroundColor.vue";
import SetColorForChartArea from "@/components/09_chart/SetColorForChartArea.vue";
import SetFont from "@/components/09_chart/SetFont.vue";
import SetFontForLegendAndDataTable from "@/components/09_chart/SetFontForLegendAndDataTable.vue";
import SetFontForTitleAndAxis from "@/components/09_chart/SetFontForTitleAndAxis.vue";
import SetLegendBackgroundColor from "@/components/09_chart/SetLegendBackgroundColor.vue";
import SetNumberFormatOfTrendline from "@/components/09_chart/SetNumberFormatOfTrendline.vue";
import ShowLeaderLine from "@/components/09_chart/ShowLeaderLine.vue";
import SparkLine from "@/components/09_chart/SparkLine.vue";
import StackedColumn from "@/components/09_chart/StackedColumn.vue";

// 10_shapes
import AddArrowLineToExcelFile from "@/components/10_shapes/AddArrowLineToExcelFile.vue";
import AddLineShape from "@/components/10_shapes/AddLineShape.vue";
import AddOvalShape from "@/components/10_shapes/AddOvalShape.vue";
import AddRectangleShape from "@/components/10_shapes/AddRectangleShape.vue";
import AddSpinnerControl from "@/components/10_shapes/AddSpinnerControl.vue";
import AdjustArrowPolylinePosition from "@/components/10_shapes/AdjustArrowPolylinePosition.vue";
import CopyShapes from "@/components/10_shapes/CopyShapes.vue";
import DeleteAllShapes from "@/components/10_shapes/DeleteAllShapes.vue";
import DeleteParticularShape from "@/components/10_shapes/DeleteParticularShape.vue";
import DrawOneLineThroughTwoPoints from "@/components/10_shapes/DrawOneLineThroughTwoPoints.vue";
import ExtractTextImageFromShape from "@/components/10_shapes/ExtractTextImageFromShape.vue";
import GetShapeLinkedCellRange from "@/components/10_shapes/GetShapeLinkedCellRange.vue";
import HideOrUnhideShape from "@/components/10_shapes/HideOrUnhideShape.vue";
import InsertShapesToExcelSheet from "@/components/10_shapes/InsertShapesToExcelSheet.vue";
import ModifyShadowStyleForShape from "@/components/10_shapes/ModifyShadowStyleForShape.vue";
import SetShadowStyleForShape from "@/components/10_shapes/SetShadowStyleForShape.vue";
import SetShapeOrder from "@/components/10_shapes/SetShapeOrder.vue";
import ShapeToImage from "@/components/10_shapes/ShapeToImage.vue";
import TillPicAsTextureInShape from "@/components/10_shapes/TillPicAsTextureInShape.vue";

//11_formatting
import ApplyBuiltInStyles from "@/components/11_formatting/ApplyBuiltInStyles.vue";
import ApplyColorScalesToDataRange from "@/components/11_formatting/ApplyColorScalesToDataRange.vue";
import ApplyConditionalFormatting from "@/components/11_formatting/ApplyConditionalFormatting.vue";
import ApplyDataBarsToCellRange from "@/components/11_formatting/ApplyDataBarsToCellRange.vue";
import ApplyGradientFillEffects from "@/components/11_formatting/ApplyGradientFillEffects.vue";
import ApplyIconSetsToCellRange from "@/components/11_formatting/ApplyIconSetsToCellRange.vue";
import ColorsAndPalette from "@/components/11_formatting/ColorsAndPalette.vue";
import ConditionalFormatRuntime from "@/components/11_formatting/ConditionalFormatRuntime.vue";
import ConditionallyFormatDate from "@/components/11_formatting/ConditionallyFormatDate.vue";
import CreateFormulaConditionalFormat from "@/components/11_formatting/CreateFormulaConditionalFormat.vue";
import FontStyles from "@/components/11_formatting/FontStyles.vue";
import ForegroundAndBackground from "@/components/11_formatting/ForegroundAndBackground.vue";
import FormatAColumn from "@/components/11_formatting/FormatAColumn.vue";
import FormatARow from "@/components/11_formatting/FormatARow.vue";
import FormatCellsWithStyle from "@/components/11_formatting/FormatCellsWithStyle.vue";
import GetColorArgbData from "@/components/11_formatting/GetColorArgbData.vue";
import GetStyleSetStyle from "@/components/11_formatting/GetStyleSetStyle.vue";
import HighlightAverageValues from "@/components/11_formatting/HighlightAverageValues.vue";
import HighlightDuplicateUniqueValues from "@/components/11_formatting/HighlightDuplicateUniqueValues.vue";
import HighlightRankedValues from "@/components/11_formatting/HighlightRankedValues.vue";
import Indentation from "@/components/11_formatting/Indentation.vue";
import Interior from "@/components/11_formatting/Interior.vue";
import MakeCellActive from "@/components/11_formatting/MakeCellActive.vue";
import NumberStyles from "@/components/11_formatting/NumberStyles.vue";
import SetBorder from "@/components/11_formatting/SetBorder.vue";
import SetBorderToDataBar from "@/components/11_formatting/SetBorderToDataBar.vue";
import SetConditionalFormatFormula from "@/components/11_formatting/SetConditionalFormatFormula.vue";
import SetRowColorByConditionalFormat from "@/components/11_formatting/SetRowColorByConditionalFormat.vue";
import SetTrafficLightsIcons from "@/components/11_formatting/SetTrafficLightsIcons.vue";
import SimpleConditionalFormatting from "@/components/11_formatting/SimpleConditionalFormatting.vue";
import TextAlign from "@/components/11_formatting/TextAlign.vue";
import TextDirection from "@/components/11_formatting/TextDirection.vue";
import UsingStyleObject from "@/components/11_formatting/UsingStyleObject.vue";
import UsePredefinedStyles from "@/components/11_formatting/UsePredefinedStyles.vue";
import VariousConditionalFormatting from "@/components/11_formatting/VariousConditionalFormatting.vue";

//12_formulas
import InsertFormulaWithNamedRange from "@/components/12_formulas/InsertFormulaWithNamedRange.vue";
import ReadFormulas from "@/components/12_formulas/ReadFormulas.vue";
import RegisterAddInFunction from "@/components/12_formulas/RegisterAddInFunction.vue";
import RemoveFormulasButKeepValues from "@/components/12_formulas/RemoveFormulasButKeepValues.vue";
import SubTotalFormula from "@/components/12_formulas/SubTotalFormula.vue";
import UseArrayFormulas from "@/components/12_formulas/UseArrayFormulas.vue";
import UseArrayR1C1Formula from "@/components/12_formulas/UseArrayR1C1Formula.vue";
import UseR1C1Formula from "@/components/12_formulas/UseR1C1Formula.vue";
import WriteFormulas from "@/components/12_formulas/WriteFormulas.vue";

//13_headerFooter
import ChangeFontAndSize from "@/components/13_headerFooter/ChangeFontAndSize.vue";
import DifferentHeaderFooter from "@/components/13_headerFooter/DifferentHeaderFooter.vue";
import DifferentHeaderFooterOnFirstPage from "@/components/13_headerFooter/DifferentHeaderFooterOnFirstPage.vue";
import ImageHeaderFooter from "@/components/13_headerFooter/ImageHeaderFooter.vue";
import SetHeaderFooter from "@/components/13_headerFooter/SetHeaderFooter.vue";

//14_hyperlinks
import AddHyperlinkToText from "@/components/14_hyperlinks/AddHyperlinkToText.vue";
import AddImageHyperlink from "@/components/14_hyperlinks/AddImageHyperlink.vue";
import GetHyperLinkType from "@/components/14_hyperlinks/GetHyperLinkType.vue";
import LinkToOtherSheetCell from "@/components/14_hyperlinks/LinkToOtherSheetCell.vue";
import ModifyHyperlink from "@/components/14_hyperlinks/ModifyHyperlink.vue";
import ReadHyperlinks from "@/components/14_hyperlinks/ReadHyperlinks.vue";
import RemoveHyperlinks from "@/components/14_hyperlinks/RemoveHyperlinks.vue";
import RetrieveExternalFileHyperlinks from "@/components/14_hyperlinks/RetrieveExternalFileHyperlinks.vue";
import WriteHyperlinks from "@/components/14_hyperlinks/WriteHyperlinks.vue";
import LinkToExternalFile from "@/components/14_hyperlinks/LinkToExternalFile.vue";

//15_markerDesigner
import AddVariableArray from "@/components/15_markerDesigner/AddVariableArray.vue";

//16_namedranges
import FormatNamedRangeCells from "@/components/16_namedranges/FormatNamedRangeCells.vue";
import GetAllNamedRange from "@/components/16_namedranges/GetAllNamedRange.vue";
import GetNamedRangeAddress from "@/components/16_namedranges/GetNamedRangeAddress.vue";
import GetSpecificNamedRange from "@/components/16_namedranges/GetSpecificNamedRange.vue";
import MergeNamedRangeCells from "@/components/16_namedranges/MergeNamedRangeCells.vue";
import NamedRanges from "@/components/16_namedranges/NamedRanges.vue";
import RemoveNamedRange from "@/components/16_namedranges/RemoveNamedRange.vue";
import RenameNamedRange from "@/components/16_namedranges/RenameNamedRange.vue";
import ScopedNamedRange from "@/components/16_namedranges/ScopedNamedRange.vue";
import SetFormulaWithNamedRange from "@/components/16_namedranges/SetFormulaWithNamedRange.vue";

//17_oleobjects
import ExtractOLEObjects from "@/components/17_oleobjects/ExtractOLEObjects.vue";
import InsertOLEObjects from "@/components/17_oleobjects/InsertOLEObjects.vue";
import InsertWavFileOLEObject from "@/components/17_oleobjects/InsertWavFileOLEObject.vue";

//18_pagesetup
import GetExcelPaperDimensions from "@/components/18_pagesetup/GetExcelPaperDimensions.vue";
import GetExcelVersion from "@/components/18_pagesetup/GetExcelVersion.vue";
import SetExcelPageOrderType from "@/components/18_pagesetup/SetExcelPageOrderType.vue";
import SetExcelPaperSize from "@/components/18_pagesetup/SetExcelPaperSize.vue";
import SetFirstPageNumber from "@/components/18_pagesetup/SetFirstPageNumber.vue";
import SetHeaderAndFooterMargins from "@/components/18_pagesetup/SetHeaderAndFooterMargins.vue";
import SetMarginsOfExcelSheet from "@/components/18_pagesetup/SetMarginsOfExcelSheet.vue";
import SetOtherPrintingOptions from "@/components/18_pagesetup/SetOtherPrintingOptions.vue";
import SetPageOrientation from "@/components/18_pagesetup/SetPageOrientation.vue";
import SetPrintAreaOfXlsFile from "@/components/18_pagesetup/SetPrintAreaOfXlsFile.vue";
import SetPrintQualityOfXlsFile from "@/components/18_pagesetup/SetPrintQualityOfXlsFile.vue";
import SetPrintTitleOfXlsFile from "@/components/18_pagesetup/SetPrintTitleOfXlsFile.vue";
import SetSheetFitToPageProperty from "@/components/18_pagesetup/SetSheetFitToPageProperty.vue";
import SetXlsSheetCenterOnPage from "@/components/18_pagesetup/SetXlsSheetCenterOnPage.vue";

//19_pivottables
import ChangeDataSource from "@/components/19_pivottables/ChangeDataSource.vue";
import ClearPivotFields from "@/components/19_pivottables/ClearPivotFields.vue";
import ConsolidationFunctions from "@/components/19_pivottables/ConsolidationFunctions.vue";
import CreatePivotTable from "@/components/19_pivottables/CreatePivotTable.vue";
import DisablePivotTableRibbon from "@/components/19_pivottables/DisablePivotTableRibbon.vue";
import ExpandOrCollapseRows from "@/components/19_pivottables/ExpandOrCollapseRows.vue";
import FormatDataField from "@/components/19_pivottables/FormatDataField.vue";
import FormattingAppearance from "@/components/19_pivottables/FormattingAppearance.vue";
import GetPivotTableRefreshedInfo from "@/components/19_pivottables/GetPivotTableRefreshedInfo.vue";
import RefreshPivotTable from "@/components/19_pivottables/RefreshPivotTable.vue";
import RepeatItemLabels from "@/components/19_pivottables/RepeatItemLabels.vue";
import SetFormatOptions from "@/components/19_pivottables/SetFormatOptions.vue";
import SetPivotFieldFormat from "@/components/19_pivottables/SetPivotFieldFormat.vue";
import ShowDataFieldInRow from "@/components/19_pivottables/ShowDataFieldInRow.vue";
import ShowSubTotals from "@/components/19_pivottables/ShowSubTotals.vue";
import SortPivotTable from "@/components/19_pivottables/SortPivotTable.vue";
import UpdateDataSource from "@/components/19_pivottables/UpdateDataSource.vue";

//20_security
import DetectProtection from "@/components/20_security/DetectProtection.vue";
import HideFormulas from "@/components/20_security/HideFormulas.vue";
import LockSpecificCellInNewExcel from "@/components/20_security/LockSpecificCellInNewExcel.vue";
import LockSpecificColumnInNewExcel from "@/components/20_security/LockSpecificColumnInNewExcel.vue";
import LockSpecificRowInNewExcel from "@/components/20_security/LockSpecificRowInNewExcel.vue";
import ProtectCell from "@/components/20_security/ProtectCell.vue";
import ProtectWithEditableRange from "@/components/20_security/ProtectWithEditableRange.vue";
import ProtectWorkbook from "@/components/20_security/ProtectWorkbook.vue";
import RemoveAllDigitalSignatures from "@/components/20_security/RemoveAllDigitalSignatures.vue";
import UnlockProtectSheet from "@/components/20_security/UnlockProtectSheet.vue";
import UnlockSimpleSheet from "@/components/20_security/UnlockSimpleSheet.vue";

//21_textboxes
import ExtractTextFromATextbox from "@/components/21_textBoxes/ExtractTextFromATextbox.vue";
import GetTextBoxByName from "@/components/21_textBoxes/GetTextBoxByName.vue";
import ManipulateTextBox from "@/components/21_textBoxes/ManipulateTextBox.vue";
import RemoveBorderlineOfTextbox from "@/components/21_textBoxes/RemoveBorderlineOfTextbox.vue";
import ReplaceTextInTextBox from "@/components/21_textBoxes/ReplaceTextInTextBox.vue";
import SetFontAndBackground from "@/components/21_textBoxes/SetFontAndBackground.vue";
import SetInternalMarginOfTextbox from "@/components/21_textBoxes/SetInternalMarginOfTextbox.vue";
import TextBoxWithWrapText from "@/components/21_textBoxes/TextBoxWithWrapText.vue";

//22_worksheets
import ActivateWorksheet from "@/components/22_worksheets/ActivateWorksheet.vue";
import AddPageBreakInXlsFile from "@/components/22_worksheets/AddPageBreakInXlsFile.vue";
import AddWorksheet from "@/components/22_worksheets/AddWorksheet.vue";
import ApplyStyleToWorksheet from "@/components/22_worksheets/ApplyStyleToWorksheet.vue";
import CheckDialogSheetInXlsFile from "@/components/22_worksheets/CheckDialogSheetInXlsFile.vue";
import CopySheetToAnotherXlsFile from "@/components/22_worksheets/CopySheetToAnotherXlsFile.vue";
import CopySheetWithinWorkbook from "@/components/22_worksheets/CopySheetWithinWorkbook.vue";
import CopyVisibleSheets from "@/components/22_worksheets/CopyVisibleSheets.vue";
import CopyWorksheet from "@/components/22_worksheets/CopyWorksheet.vue";
import DetectEmptyWorksheet from "@/components/22_worksheets/DetectEmptyWorksheet.vue";
import FillDataInWorksheet from "@/components/22_worksheets/FillDataInWorksheet.vue";
import FreezePanes from "@/components/22_worksheets/FreezePanes.vue";
import GetFreezePaneRange from "@/components/22_worksheets/GetFreezePaneRange.vue";
import GetListOfFontsUsed from "@/components/22_worksheets/GetListOfFontsUsed.vue";
import GetPaperSize from "@/components/22_worksheets/GetPaperSize.vue";
import GetWorksheetNames from "@/components/22_worksheets/GetWorksheetNames.vue";
import HideOrShowWorksheet from "@/components/22_worksheets/HideOrShowWorksheet.vue";
import HideTab from "@/components/22_worksheets/HideTab.vue";
import HideZeroValues from "@/components/22_worksheets/HideZeroValues.vue";
import LinkToContentProperty from "@/components/22_worksheets/LinkToContentProperty.vue";
import MoveWorksheet from "@/components/22_worksheets/MoveWorksheet.vue";
import PageBreakPreview from "@/components/22_worksheets/PageBreakPreview.vue";
import RemovePageBreak from "@/components/22_worksheets/RemovePageBreak.vue";
import RemoveWorksheet from "@/components/22_worksheets/RemoveWorksheet.vue";
import SetPageBreak from "@/components/22_worksheets/SetPageBreak.vue";
import SetTabColor from "@/components/22_worksheets/SetTabColor.vue";
import SetViewMode from "@/components/22_worksheets/SetViewMode.vue";
import ShowOrHideGridLine from "@/components/22_worksheets/ShowOrHideGridLine.vue";
import ShowTab from "@/components/22_worksheets/ShowTab.vue";
import SplitWorksheetIntoPanes from "@/components/22_worksheets/SplitWorksheetIntoPanes.vue";
import UnfreezeExcelPanes from "@/components/22_worksheets/UnfreezeExcelPanes.vue";
import VerifyProtectedWorksheet from "@/components/22_worksheets/VerifyProtectedWorksheet.vue";
import ZoomFactor from "@/components/22_worksheets/ZoomFactor.vue";

//23_workbook
import AccessDocumentProperties from "@/components/23_workbook/AccessDocumentProperties.vue";
import AddCustomProperties from "@/components/23_workbook/AddCustomProperties.vue";
import DecryptWorkbook from "@/components/23_workbook/DecryptWorkbook.vue";
import DetectExcelVersion from "@/components/23_workbook/DetectExcelVersion.vue";
import DetectVBAMacros from "@/components/23_workbook/DetectVBAMacros.vue";
import EncryptWorkbook from "@/components/23_workbook/EncryptWorkbook.vue";
import GetProperties from "@/components/23_workbook/GetProperties.vue";
import HideWindow from "@/components/23_workbook/HideWindow.vue";
import LoadAndSaveFileWithMacro from "@/components/23_workbook/LoadAndSaveFileWithMacro.vue";
import MergeExcelFiles from "@/components/23_workbook/MergeExcelFiles.vue";
import OpenEncryptedFile from "@/components/23_workbook/OpenEncryptedFile.vue";
import OpenFiles from "@/components/23_workbook/OpenFiles.vue";
import ReadStream from "@/components/23_workbook/ReadStream.vue";
import RemoveCustomProperties from "@/components/23_workbook/RemoveCustomProperties.vue";
import SaveFiles from "@/components/23_workbook/SaveFiles.vue";
import SaveStream from "@/components/23_workbook/SaveStream.vue";
import SetExcelCalculationMode from "@/components/23_workbook/SetExcelCalculationMode.vue";
import SetMargins from "@/components/23_workbook/SetMargins.vue";
import SetTheme from "@/components/23_workbook/SetTheme.vue";

// 24_trackChanges
import AcceptOrRejectTrackedChanges from "@/components/24_trackChanges/AcceptOrRejectTrackedChanges.vue";

const filePageId = ref<string>("0");
const handleChildID = (id: any) => {
  filePageId.value = id;
};
</script>

<style lang="scss" scoped>
* {
  padding: 0%;
  margin: 0%;
}
button {
  padding: 5px 10px;
}
.worker_page {
  width: 100%;
  height: 100vh;
  display: flex;
  padding: 2rem;
  .tree_item {
    width: 20%;
    border-right: 1px solid #eeeeee;
  }
  .file_item {
    display: flex;
    flex: 1;
    justify-content: center;
    align-items: center;
    .file-reader div:first-of-type {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 20px;
      span {
        font-size: 18px;
        font-weight: 500;
      }
    }
    .non-centered {
      align-items: flex-start;
    }
  }
}

pre {
  text-align: left;
  background-color: #f8f8f8;
  padding: 10px;
  border: 1px solid #ccc;
}
</style>
