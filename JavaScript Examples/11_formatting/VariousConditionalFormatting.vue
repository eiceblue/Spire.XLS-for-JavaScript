<template>
  <span>Click the following button to add conditional formatting in a new excel</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
		wasmModule=window.wasmModule;
      if (wasmModule) {
        // Load font
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

		function AddConditionalFormattingForNewSheet(sheet) {
			AddDefaultIconSet(sheet);
			AddIconSet2(sheet);
			AddIconSet3(sheet);
			AddIconSet4(sheet);
			AddIconSet5(sheet);
			AddIconSet6(sheet);
			AddIconSet7(sheet);
			AddIconSet8(sheet);
			AddIconSet9(sheet);
			AddIconSet10(sheet);
			AddIconSet11(sheet);
			AddIconSet12(sheet);
			AddIconSet13(sheet);
			AddIconSet14(sheet);
			AddIconSet15(sheet);
			AddIconSet16(sheet);
			AddIconSet17(sheet);
			AddIconSet18(sheet);
			AddDefaultColorScale(sheet);
			Add3ColorScale(sheet);
			Add2ColorScale(sheet);
			AddAboveAverage(sheet);
			AddAboveAverage2(sheet);
			AddAboveAverage3(sheet);
			AddTop10_1(sheet);
			AddTop10_2(sheet);
			AddTop10_3(sheet);
			AddTop10_4(sheet);
			AddDataBar1(sheet);
			AddDataBar2(sheet);
			AddContainsText(sheet);
			AddNotContainsText(sheet);
			AddContainsBlank(sheet);
			AddNotContainsBlank(sheet);
			AddBeginWith(sheet);
			AddEndWith(sheet);
			AddContainsError(sheet);
			AddNotContainsError(sheet);
			AddDuplicate(sheet);
			AddUnique(sheet);
			AddTimePeriod_1(sheet);
			AddTimePeriod_2(sheet);
			AddTimePeriod_3(sheet);
			AddTimePeriod_4(sheet);
			AddTimePeriod_5(sheet);
			AddTimePeriod_6(sheet);
			AddTimePeriod_7(sheet);
			AddTimePeriod_8(sheet);
			AddTimePeriod_9(sheet);
			AddTimePeriod_10(sheet);
			sheet.AllocatedRange.ColumnWidth = 15;
			sheet.AllocatedRange.AutoFitRows();
		}

		//This method implements the IconSet conditional formatting type with ThreeArrows colored attribute.
	function AddIconSet2(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M1:O2"));
			sheet.Range.get("M1:O2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M1:O2").Style.Color = wasmModule.Color.get_AliceBlue();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeArrows;
			sheet.Range.get("M1").Text = "ThreeArrows";
			sheet.Range.get("N1").NumberValue = 15;
			sheet.Range.get("O1").NumberValue = 18;
			sheet.Range.get("M2").NumberValue = 14;
			sheet.Range.get("N2").NumberValue = 17;
			sheet.Range.get("O2").NumberValue = 20;
		}

		//This method implements the IconSet conditional formatting type with FourArrows colored attribute.
	function AddIconSet3(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M3:O4"));
			sheet.Range.get("M3:O4").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M3:O4").Style.Color = wasmModule.Color.get_AntiqueWhite();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FourArrows;
			sheet.Range.get("M3").Text = "FourArrows";
			sheet.Range.get("N3").NumberValue = 17;
			sheet.Range.get("O3").NumberValue = 20;
			sheet.Range.get("M4").NumberValue = 16;
			sheet.Range.get("N4").NumberValue = 19;
			sheet.Range.get("O4").NumberValue = 22;
		}

		//This method implements the IconSet conditional formatting type with FiveArrows colored attribute.
	function AddIconSet4(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M5:O6"));
			sheet.Range.get("M5:O6").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M5:O6").Style.Color = wasmModule.Color.get_Aqua();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FiveArrows;
			sheet.Range.get("M5").Text = "FiveArrows";
			sheet.Range.get("N5").NumberValue = 17;
			sheet.Range.get("O5").NumberValue = 20;
			sheet.Range.get("M6").NumberValue = 16;
			sheet.Range.get("N6").NumberValue = 19;
			sheet.Range.get("O6").NumberValue = 22;
		}

		//This method implements the IconSet conditional formatting type with ThreeArrowsGray colored attribute.
	function AddIconSet5(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M7:O8"));
			sheet.Range.get("M7:O8").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M7:O8").Style.Color = wasmModule.Color.get_Aquamarine();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeArrowsGray;
			sheet.Range.get("M7").Text = "ThreeArrowsGray";
			sheet.Range.get("N7").NumberValue = 21;
			sheet.Range.get("O7").NumberValue = 24;
			sheet.Range.get("M8").NumberValue = 20;
			sheet.Range.get("N8").NumberValue = 23;
			sheet.Range.get("O8").NumberValue = 26;
		}

		//This method implements the IconSet conditional formatting type with FourArrowsGray colored attribute.
	function AddIconSet6(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M9:O10"));
			sheet.Range.get("M9:O10").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M9:O10").Style.Color = wasmModule.Color.get_Azure();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FourArrowsGray;
			sheet.Range.get("M9").Text = "FourArrowsGray";
			sheet.Range.get("N9").NumberValue = 23;
			sheet.Range.get("O9").NumberValue = 26;
			sheet.Range.get("M10").NumberValue = 22;
			sheet.Range.get("N10").NumberValue = 25;
			sheet.Range.get("O10").NumberValue = 28;
		}

		//This method implements the IconSet conditional formatting type with FiveArrowsGray colored attribute.
	function AddIconSet7(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M11:O12"));
			sheet.Range.get("M11:O12").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M11:O12").Style.Color = wasmModule.Color.get_Beige();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FiveArrowsGray;
			sheet.Range.get("M11").Text = "FiveArrowsGray";
			sheet.Range.get("N11").NumberValue = 25;
			sheet.Range.get("O11").NumberValue = 28;
			sheet.Range.get("M12").NumberValue = 24;
			sheet.Range.get("N12").NumberValue = 27;
			sheet.Range.get("O12").NumberValue = 30;
		}

		//This method implements the IconSet conditional formatting type with ThreeFlags attribute.
	function AddIconSet8(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M13:O14"));
			sheet.Range.get("M13:O14").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M13:O14").Style.Color = wasmModule.Color.get_Bisque();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeFlags;
			sheet.Range.get("M13").Text = "ThreeFlags";
			sheet.Range.get("N13").NumberValue = 27;
			sheet.Range.get("O13").NumberValue = 30;
			sheet.Range.get("M14").NumberValue = 26;
			sheet.Range.get("N14").NumberValue = 29;
			sheet.Range.get("O14").NumberValue = 32;
		}

		//This method implements the IconSet conditional formatting type with FiveQuarters attribute.
	function AddIconSet9(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M15:O16"));
			sheet.Range.get("M15:O16").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M15:O16").Style.Color = wasmModule.Color.get_BlanchedAlmond();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FiveQuarters;
			sheet.Range.get("M15").Text = "FiveQuarters";
			sheet.Range.get("N15").NumberValue = 29;
			sheet.Range.get("O15").NumberValue = 32;
			sheet.Range.get("M16").NumberValue = 28;
			sheet.Range.get("N16").NumberValue = 31;
			sheet.Range.get("O16").NumberValue = 34;
		}

		//This method implements the IconSet conditional formatting type with FourRating attribute.
	function AddIconSet10(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M17:O18"));
			sheet.Range.get("M17:O18").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M17:O18").Style.Color = wasmModule.Color.get_LightBlue();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FourRating;
			sheet.Range.get("M17").Text = "FourRating";
			sheet.Range.get("N17").NumberValue = 31;
			sheet.Range.get("O17").NumberValue = 34;
			sheet.Range.get("M18").NumberValue = 30;
			sheet.Range.get("N18").NumberValue = 33;
			sheet.Range.get("O18").NumberValue = 36;
		}

		//This method implements the IconSet conditional formatting type with FiveRating attribute.
	function AddIconSet11(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M19:O20"));
			sheet.Range.get("M19:O20").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M19:O20").Style.Color = wasmModule.Color.get_BlueViolet();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FiveRating;
			sheet.Range.get("M19").Text = "FiveRating";
			sheet.Range.get("N19").NumberValue = 33;
			sheet.Range.get("O19").NumberValue = 36;
			sheet.Range.get("M20").NumberValue = 32;
			sheet.Range.get("N20").NumberValue = 35;
			sheet.Range.get("O20").NumberValue = 38;
		}

		//This method implements the IconSet conditional formatting type with FourRedToBlack attribute.
	//This method implements the IconSet conditional formatting type with FourRedToBlack attribute.
	function AddIconSet12(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M21:O22"));
			sheet.Range.get("M21:O22").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M21:O22").Style.Color = wasmModule.Color.get_Brown();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FourRedToBlack;
			sheet.Range.get("M21").Text = "FourRedToBlack";
			sheet.Range.get("N21").NumberValue = 35;
			sheet.Range.get("O21").NumberValue = 38;
			sheet.Range.get("M22").NumberValue = 34;
			sheet.Range.get("N22").NumberValue = 37;
			sheet.Range.get("O22").NumberValue = 40;
		}

	//This method implements the IconSet conditional formatting type with ThreeSigns attribute.
	function AddIconSet13(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M23:O24"));
			sheet.Range.get("M23:O24").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M23:O24").Style.Color = wasmModule.Color.get_BurlyWood();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeSigns;
			sheet.Range.get("M23").Text = "ThreeSigns";
			sheet.Range.get("N23").NumberValue = 37;
			sheet.Range.get("O23").NumberValue = 40;
			sheet.Range.get("M24").NumberValue = 36;
			sheet.Range.get("N24").NumberValue = 39;
			sheet.Range.get("O24").NumberValue = 42;
		}

	//This method implements the IconSet conditional formatting type with ThreeSymbols attribute.
	function AddIconSet14(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M25:O26"));
			sheet.Range.get("M25:O26").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M25:O26").Style.Color = wasmModule.Color.get_CadetBlue();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeSymbols;
			sheet.Range.get("M25").Text = "ThreeSymbols";
			sheet.Range.get("N25").NumberValue = 39;
			sheet.Range.get("O25").NumberValue = 42;
			sheet.Range.get("M26").NumberValue = 38;
			sheet.Range.get("N26").NumberValue = 41;
			sheet.Range.get("O26").NumberValue = 44;
		}

	//This method implements the IconSet conditional formatting type with ThreeSymbols2 attribute.
	function AddIconSet15(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M27:O28"));
			sheet.Range.get("M27:O28").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M27:O28").Style.Color = wasmModule.Color.get_Chartreuse();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeSymbols2;
			sheet.Range.get("M27").Text = "ThreeSymbols2";
			sheet.Range.get("N27").NumberValue = 41;
			sheet.Range.get("O27").NumberValue = 44;
			sheet.Range.get("M28").NumberValue = 40;
			sheet.Range.get("N28").NumberValue = 43;
			sheet.Range.get("O28").NumberValue = 46;
		}

	//This method implements the IconSet conditional formatting type with ThreeTrafficLights1 attribute.
	function AddIconSet16(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M29:O30"));
			sheet.Range.get("M29:O30").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M29:O30").Style.Color = wasmModule.Color.get_Chocolate();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights1;
			sheet.Range.get("M29").Text = "ThreeTrafficLights1";
			sheet.Range.get("N29").NumberValue = 43;
			sheet.Range.get("O29").NumberValue = 46;
			sheet.Range.get("M30").NumberValue = 42;
			sheet.Range.get("N30").NumberValue = 45;
			sheet.Range.get("O30").NumberValue = 48;
		}

	//This method implements the IconSet conditional formatting type with ThreeTrafficLights2 attribute.
	function AddIconSet17(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M31:O32"));
			sheet.Range.get("M31:O32").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M31:O32").Style.Color = wasmModule.Color.get_Coral();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.ThreeTrafficLights2;
			sheet.Range.get("M31").Text = "ThreeTrafficLights2";
			sheet.Range.get("N31").NumberValue = 45;
			sheet.Range.get("O31").NumberValue = 48;
			sheet.Range.get("M32").NumberValue = 44;
			sheet.Range.get("N32").NumberValue = 47;
			sheet.Range.get("O32").NumberValue = 50;
		}

	//This method implements the IconSet conditional formatting type with FourTrafficLights attribute.
	function AddIconSet18(sheet) {
			let xcfs = sheet.ConditionalFormats.Add();
			xcfs.AddRange(sheet.Range.get("M33:O35"));
			sheet.Range.get("M33:O35").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("M33:O35").Style.Color = wasmModule.Color.get_CornflowerBlue();
			let cf = xcfs.AddCondition();
			cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
			cf.IconSet.IconSetType = wasmModule.IconSetType.FourTrafficLights;
			sheet.Range.get("M33").Text = "FourTrafficLights";
			sheet.Range.get("N33").NumberValue = 48;
			sheet.Range.get("O33").NumberValue = 52;
			sheet.Range.get("M34").NumberValue = 46;
			sheet.Range.get("N34").NumberValue = 50;
			sheet.Range.get("O34").NumberValue = 54;
			sheet.Range.get("M35").NumberValue = 48;
			sheet.Range.get("N35").NumberValue = 52;
			sheet.Range.get("O35").NumberValue = 56;
		}

	//This method implements the TimePeriod conditional formatting type with Yesterday attribute.
	function AddTimePeriod_10(sheet) {
			let conds = sheet.ConditionalFormats.Add();
			conds.AddRange(sheet.Range.get("I19:K20"));
			sheet.Range.get("I19:K20").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("I19:K20").Style.Color = wasmModule.Color.get_MediumSeaGreen();
			let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.Yesterday);
			cf.FillPattern = wasmModule.ExcelPatternType.Solid;
			cf.BackColor = wasmModule.Color.get_Pink();
			let c = sheet.Range.get("I19");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(-2);

			c = sheet.Range.get("J19");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(-1);

			c = sheet.Range.get("K19");
			c.Value2 = wasmModule.DateTime.get_Now();

			c = sheet.Range.get("I20");
			c.Text = "Yesterday";

			c = sheet.Range.get("J20");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(1);

			c = sheet.Range.get("K20");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(2);
		}

	//This method implements the TimePeriod conditional formatting type with Tomorrow attribute.
	function AddTimePeriod_9(sheet) {
			let conds = sheet.ConditionalFormats.Add();
			conds.AddRange(sheet.Range.get("I17:K18"));
			sheet.Range.get("I17:K18").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("I17:K18").Style.Color = wasmModule.Color.get_MediumPurple();
			let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.Tomorrow);
			cf.FillPattern = wasmModule.ExcelPatternType.Solid;
			cf.BackColor = wasmModule.Color.get_Pink();
			let c = sheet.Range.get("I17");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(-2);

			c = sheet.Range.get("J17");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(-1);

			c = sheet.Range.get("K17");
			c.Value2 = wasmModule.DateTime.get_Now();

			c = sheet.Range.get("I18");
			c.Text = "Tomorrow";

			c = sheet.Range.get("J18");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(1);

			c = sheet.Range.get("K18");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(2);
		}

	//This method implements the TimePeriod conditional formatting type with ThisWeek attribute.
	function AddTimePeriod_8(sheet) {
			let conds = sheet.ConditionalFormats.Add();
			conds.AddRange(sheet.Range.get("I15:K16"));
			sheet.Range.get("I15:K16").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
			sheet.Range.get("I15:K16").Style.Color = wasmModule.Color.get_MediumOrchid();
			let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.ThisWeek);
			cf.FillPattern = wasmModule.ExcelPatternType.Solid;
			cf.BackColor = wasmModule.Color.get_Pink();
			let c = sheet.Range.get("I15");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(-2);

			c = sheet.Range.get("J15");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(-1);

			c = sheet.Range.get("K15");
			c.Value2 = wasmModule.DateTime.get_Now();

			c = sheet.Range.get("I16");
			c.Text = "ThisWeek";

			c = sheet.Range.get("J16");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(2);

			c = sheet.Range.get("K16");
			c.Value2 = wasmModule.DateTime.get_Now().AddDays(1);
		}


	// This method implements the TimePeriod conditional formatting type with Today attribute.
	function AddTimePeriod_1(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I1:K2"));
		sheet.Range.get("I1:K2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I1:K2").Style.Color = wasmModule.Color.get_LightSlateGray();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.Today);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I1");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-8);
		c = sheet.Range.get("J1");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-7);
		c = sheet.Range.get("K1");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I2");
		c.Text = "Today";
		c = sheet.Range.get("J2");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(3);
		c = sheet.Range.get("K2");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(2);
	}

	// This method implements the TimePeriod conditional formatting type with Last7Days attribute.
	function AddTimePeriod_2(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I3:K4"));
		sheet.Range.get("I3:K4").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I3:K4").Style.Color = wasmModule.Color.get_LightSkyBlue();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.Last7Days);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I3");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-8);
		c = sheet.Range.get("J3");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-7);
		c = sheet.Range.get("K3");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I4");
		c.Text = "Last7Days";
		c = sheet.Range.get("J4");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(3);
		c = sheet.Range.get("K4");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(2);
	}

	// This method implements the TimePeriod conditional formatting type with LastMonth attribute.
	function AddTimePeriod_3(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I5:K6"));
		sheet.Range.get("I5:K6").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I5:K6").Style.Color = wasmModule.Color.get_Linen();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.LastMonth);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I5");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-6);
		c = sheet.Range.get("J5");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(-1);
		c = sheet.Range.get("K5");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I6");
		c.Text = "LastMonth";
		c = sheet.Range.get("J6");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(3);
		c = sheet.Range.get("K6");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(1);
	}

	// This method implements the TimePeriod conditional formatting type with LastWeek attribute.
	function AddTimePeriod_4(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I7:K8"));
		sheet.Range.get("I7:K8").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I7:K8").Style.Color = wasmModule.Color.get_Linen();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.LastWeek);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I7");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-6);
		c = sheet.Range.get("J7");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-5);
		c = sheet.Range.get("K7");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I8");
		c.Text = "LastWeek";
		c = sheet.Range.get("J8");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(3);
		c = sheet.Range.get("K8");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(4);
	}

	// This method implements the TimePeriod conditional formatting type with NextMonth attribute.
	function AddTimePeriod_5(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I9:K10"));
		sheet.Range.get("I9:K10").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I9:K10").Style.Color = wasmModule.Color.get_Maroon();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.NextMonth);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I9");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-3);
		c = sheet.Range.get("J9");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(-1);
		c = sheet.Range.get("K9");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I10");
		c.Text = "NextMonth";
		c = sheet.Range.get("J10");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(1);
		c = sheet.Range.get("K10");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(2);
	}

	// This method implements the TimePeriod conditional formatting type with NextWeek attribute.
	function AddTimePeriod_6(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I11:K12"));
		sheet.Range.get("I11:K12").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I11:K12").Style.Color = wasmModule.Color.get_MediumAquamarine();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.NextWeek);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I11");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-3);
		c = sheet.Range.get("J11");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-2);
		c = sheet.Range.get("K11");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I12");
		c.Text = "NextWeek";
		c = sheet.Range.get("J12");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(3);
		c = sheet.Range.get("K12");
		c.Value2 = wasmModule.DateTime.get_Now().AddMonths(4);
	}

	// This method implements the TimePeriod conditional formatting type with ThisMonth attribute.
	function AddTimePeriod_7(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("I13:K14"));
		sheet.Range.get("I13:K14").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("I13:K14").Style.Color = wasmModule.Color.get_MediumBlue();
		let cf = conds.AddTimePeriodCondition(wasmModule.TimePeriodType.ThisMonth);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("I13");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-1);
		c = sheet.Range.get("J13");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(-1);
		c = sheet.Range.get("K13");
		c.Value2 = wasmModule.DateTime.get_Now();
		c = sheet.Range.get("I14");
		c.Text = "ThisMonth";
		c = sheet.Range.get("J14");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(1);
		c = sheet.Range.get("K14");
		c.Value2 = wasmModule.DateTime.get_Now().AddDays(2);
	}

	// This method implements the UniqueValues conditional formatting type.
	function AddUnique(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E21:G22"));
		sheet.Range.get("E21:G22").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E21:G22").Style.Color = wasmModule.Color.get_LightSalmon();
		let cf = conds.AddUniqueValuesCondition();
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();
		let c = sheet.Range.get("E21");
		c.Text = "aa";
		c = sheet.Range.get("F21");
		c.Text = "bb";
		c = sheet.Range.get("G21");
		c.Text = "aa";
		c = sheet.Range.get("E22");
		c.Text = "bbb";
		c = sheet.Range.get("F22");
		c.Text = "bb";
		c = sheet.Range.get("G22");
		c.Text = "ccc";
	}

	//This method implements the DuplicateValues conditional formatting type.
	function AddDuplicate(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E23:G24"));
		sheet.Range.get("E23:G24").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E23:G24").Style.Color = wasmModule.Color.get_LightSlateGray();
		let cf = conds.AddDuplicateValuesCondition();
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();
		let c = sheet.Range.get("E23");
		c.Text = "aa";
		c = sheet.Range.get("F23");
		c.Text = "bb";
		c = sheet.Range.get("G23");
		c.Text = "aa";
		c = sheet.Range.get("E24");
		c.Text = "bbb";
		c = sheet.Range.get("F24");
		c.Text = "bb";
		c = sheet.Range.get("G24");
		c.Text = "ccc";
	}


	//This method implements the NotContainsError conditional formatting type.
	function AddNotContainsError(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E19:G20"));
		sheet.Range.get("E19:G20").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E19:G20").Style.Color = wasmModule.Color.get_LightSeaGreen();
		let cf = conds.AddNotContainsErrorsCondition();
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();
		let c = sheet.Range.get("E19");
		c.Text = "aa";
		c = sheet.Range.get("F19");
		c.Text = "=Sum";
		c = sheet.Range.get("G19");
		c.Text = "aa";
		c = sheet.Range.get("E20");
		c.Text = "bbb";
		c = sheet.Range.get("F20");
		c.Text = "sss";
		c = sheet.Range.get("G20");
		c.Text = "=Max";
	}

	//This method implements the ContainsErrors conditional formatting type.
	function AddContainsError(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E17:G18"));
		sheet.Range.get("E17:G18").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E17:G18").Style.Color = wasmModule.Color.get_LightSkyBlue();
		let cf = conds.AddContainsErrorsCondition();
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();
		let c = sheet.Range.get("E17");
		c.Text = "aa";
		c = sheet.Range.get("F17");
		c.Text = "=Sum";
		c = sheet.Range.get("G17");
		c.Text = "aa";
		c = sheet.Range.get("E18");
		c.Text = "bbb";
		c = sheet.Range.get("F18");
		c.Text = "sss";
		c = sheet.Range.get("G18");
		c.Text = "=Max";
	}

	//This method implements the BeginWith conditional formatting type.
	function AddBeginWith(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E15:G16"));
		sheet.Range.get("E15:G16").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E15:G16").Style.Color = wasmModule.Color.get_LightGoldenrodYellow();
		let cf = conds.AddBeginsWithCondition("ab");
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();

		let c = sheet.Range.get("E15");
		c.Text = "aa";
		c = sheet.Range.get("F15");
		c.Text = "abc";
		c = sheet.Range.get("G15");
		c.Text = "aa";
		c = sheet.Range.get("E16");
		c.Text = "bbb";
		c = sheet.Range.get("F16");
		c.Text = "sss";
		c = sheet.Range.get("G16");
		c.Text = "abcd";
	}

	//This method implements the EndWith conditional formatting type.
	function AddEndWith(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E13:G14"));
		sheet.Range.get("E13:G14").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E13:G14").Style.Color = wasmModule.Color.get_LightGray();
		let cf = conds.AddEndsWithCondition("ab");
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();

		let c = sheet.Range.get("E13");
		c.Text = "aa";
		c = sheet.Range.get("F13");
		c.Text = "abc";
		c = sheet.Range.get("G13");
		c.Text = "aab";
		c = sheet.Range.get("E14");
		c.Text = "bbbc";
		c = sheet.Range.get("F14");
		c.Text = "sab";
		c = sheet.Range.get("G14");
		c.Text = "abcd";
	}

	//This method implements the NotContainsBlank conditional formatting type.
	function AddNotContainsBlank(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E11:G12"));
		sheet.Range.get("E11:G12").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E11:G12").Style.Color = wasmModule.Color.get_LightCoral();
		let cf = conds.AddNotContainsBlanksCondition();
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();

		let c = sheet.Range.get("E11");
		c.Text = "aa";
		c = sheet.Range.get("F11");
		c.Text = "  ";
		c = sheet.Range.get("G11");
		c.Text = "aab";
		c = sheet.Range.get("E12");
		c.Text = "abc";
		c = sheet.Range.get("F12");
		c.Text = "  ";
		c = sheet.Range.get("G12");
		c.Text = "abcd";
	}

	function AddContainsBlank(sheet) {
		const conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E9:G10"));
		sheet.Range.get("E9:G10").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E9:G10").Style.Color = wasmModule.Color.get_LightCyan();
		const cf = conds.AddContainsBlanksCondition();
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();

		sheet.Range.get("E9").Text = "aa";
		sheet.Range.get("F9").Text = "  ";
		sheet.Range.get("G9").Text = "aab";
		sheet.Range.get("E10").Text = "abc";
		sheet.Range.get("F10").Text = "dvdf";
		sheet.Range.get("G10").Text = "abcd";
	}

	function AddNotContainsText(sheet) {
		const conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E7:G8"));
		sheet.Range.get("E7:G8").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E7:G8").Style.Color = wasmModule.Color.get_LightGreen();
		const cf = conds.AddNotContainsTextCondition("abc");
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();

		sheet.Range.get("E7").Text = "aa";
		sheet.Range.get("F7").Text = "abfd";
		sheet.Range.get("G7").Text = "aab";
		sheet.Range.get("E8").Text = "abc";
		sheet.Range.get("F8").Text = "cedf";
		sheet.Range.get("G8").Text = "abcd";
	}

	function AddContainsText(sheet) {
		const conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("E5:G6"));
		sheet.Range.get("E5:G6").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E5:G6").Style.Color = wasmModule.Color.get_LightBlue();
		const cf = conds.AddContainsTextCondition("abc");
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();

		sheet.Range.get("E5").Text = "aa";
		sheet.Range.get("F5").Text = "abfd";
		sheet.Range.get("G5").Text = "aab";
		sheet.Range.get("E6").Text = "abc";
		sheet.Range.get("F6").Text = "cedf";
		sheet.Range.get("G6").Text = "abcd";
	}

	function AddDataBar2(sheet) {
		const xcfs = sheet.ConditionalFormats.Add();
		xcfs.AddRange(sheet.Range.get("E3:G4"));
		sheet.Range.get("E3:G4").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E3:G4").Style.Color = wasmModule.Color.get_LightGreen();
		const cf = xcfs.AddCondition();
		cf.FormatType = wasmModule.ConditionalFormatType.DataBar;
		cf.DataBar.BarColor = wasmModule.Color.get_Orange();
		cf.DataBar.MinPoint.Type = wasmModule.ConditionValueType.Percentile;
		cf.DataBar.MinPoint.Value = 30.78;
		cf.DataBar.ShowValue = false;

		sheet.Range.get("E3").NumberValue = 6;
		sheet.Range.get("F3").NumberValue = 9;
		sheet.Range.get("G3").NumberValue = 12;
		sheet.Range.get("E4").NumberValue = 8;
		sheet.Range.get("F4").NumberValue = 11;
		sheet.Range.get("G4").NumberValue = 14;
	}

	function AddDataBar1(sheet) {
		const xcfs = sheet.ConditionalFormats.Add();
		xcfs.AddRange(sheet.Range.get("E1:G2"));
		sheet.Range.get("E1:G2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("E1:G2").Style.Color = wasmModule.Color.get_YellowGreen();
		const cf = xcfs.AddCondition();
		cf.FormatType = wasmModule.ConditionalFormatType.DataBar;
		cf.DataBar.BarColor = wasmModule.Color.get_Blue();
		cf.DataBar.MinPoint.Type = wasmModule.ConditionValueType.Percent;
		cf.DataBar.ShowValue = true;

		sheet.Range.get("E1").NumberValue = 4;
		sheet.Range.get("F1").NumberValue = 7;
		sheet.Range.get("G1").NumberValue = 10;
		sheet.Range.get("E2").NumberValue = 6;
		sheet.Range.get("F2").NumberValue = 9;
		sheet.Range.get("G2").NumberValue = 14;
	}
	function AddDefaultIconSet(sheet) {
		const xcfs = sheet.ConditionalFormats.Add();
		xcfs.AddRange(sheet.Range.get("A1:C2"));
		sheet.Range.get("A1:C2").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A1:C2").Style.Color = wasmModule.Color.get_Yellow();
		const cf = xcfs.AddCondition();
		cf.FormatType = wasmModule.ConditionalFormatType.IconSet;
		sheet.Range.get("A1").NumberValue = 0;
		sheet.Range.get("B1").NumberValue = 3;
		sheet.Range.get("C1").NumberValue = 6;
		sheet.Range.get("A2").NumberValue = 2;
		sheet.Range.get("B2").NumberValue = 5;
		sheet.Range.get("C2").NumberValue = 8;
	}
	function AddDefaultColorScale(sheet) {
		const xcfs = sheet.ConditionalFormats.Add();
		xcfs.AddRange(sheet.Range.get("A5:C6"));
		sheet.Range.get("A5:C6").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A5:C6").Style.Color = wasmModule.Color.get_Pink();
		const cf = xcfs.AddCondition();
		cf.FormatType = wasmModule.ConditionalFormatType.ColorScale;

		sheet.Range.get("A5").NumberValue = 4;
		sheet.Range.get("B5").NumberValue = 7;
		sheet.Range.get("C5").NumberValue = 10;
		sheet.Range.get("A6").NumberValue = 6;
		sheet.Range.get("B6").NumberValue = 9;
		sheet.Range.get("C6").NumberValue = 12;
	}
	function Add3ColorScale(sheet) {
		const xcfs = sheet.ConditionalFormats.Add();
		xcfs.AddRange(sheet.Range.get("A7:C8"));
		sheet.Range.get("A7:C8").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A7:C8").Style.Color = wasmModule.Color.get_Green();
		const cf = xcfs.AddCondition();
		cf.FormatType = wasmModule.ConditionalFormatType.ColorScale;
		cf.ColorScale.MinValue.Type = wasmModule.ConditionValueType.Number;
		cf.ColorScale.MinValue.Value = 9;
		cf.ColorScale.MinColor = wasmModule.Color.get_Purple();

		sheet.Range.get("A7").NumberValue = 6;
		sheet.Range.get("B7").NumberValue = 9;
		sheet.Range.get("C7").NumberValue = 12;
		sheet.Range.get("A8").NumberValue = 8;
		sheet.Range.get("B8").NumberValue = 11;
		sheet.Range.get("C8").NumberValue = 14;
	}
	function Add2ColorScale(sheet) {
		const xcfs = sheet.ConditionalFormats.Add();
		xcfs.AddRange(sheet.Range.get("A9:C10"));
		sheet.Range.get("A9:C10").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A9:C10").Style.Color = wasmModule.Color.get_White();
		const cf = xcfs.AddCondition();
		cf.FormatType = wasmModule.ConditionalFormatType.ColorScale;
		cf.ColorScale.MinColor = wasmModule.Color.get_Gold();
		cf.ColorScale.MaxColor = wasmModule.Color.get_SkyBlue();

		sheet.Range.get("A9").NumberValue = 8;
		sheet.Range.get("B9").NumberValue = 12;
		sheet.Range.get("C9").NumberValue = 13;
		sheet.Range.get("A10").NumberValue = 10;
		sheet.Range.get("B10").NumberValue = 13;
		sheet.Range.get("C10").NumberValue = 16;
	}
	function AddAboveAverage(sheet) {
		const conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A11:C12"));
		sheet.Range.get("A11:C12").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A11:C12").Style.Color = wasmModule.Color.get_Tomato();
		const cf = conds.AddAverageCondition(wasmModule.AverageType.Above);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();

		sheet.Range.get("A11").NumberValue = 10;
		sheet.Range.get("B11").NumberValue = 13;
		sheet.Range.get("C11").NumberValue = 16;
		sheet.Range.get("A12").NumberValue = 12;
		sheet.Range.get("B12").NumberValue = 15;
		sheet.Range.get("C12").NumberValue = 18;
	}
	function AddAboveAverage2(sheet) {
		const conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A13:C14"));
		sheet.Range.get("A13:C14").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A13:C14").Style.Color = wasmModule.Color.get_LightPink();
		const cf = conds.AddAverageCondition(wasmModule.AverageType.BelowEqual);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_LightSkyBlue();

		sheet.Range.get("A13").NumberValue = 12;
		sheet.Range.get("B13").NumberValue = 15;
		sheet.Range.get("C13").NumberValue = 18;
		sheet.Range.get("A14").NumberValue = 14;
		sheet.Range.get("B14").NumberValue = 17;
		sheet.Range.get("C14").NumberValue = 20;
	}

	function AddAboveAverage3(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A15:C16"));
		sheet.Range.get("A15:C16").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A15:C16").Style.Color = wasmModule.Color.get_LightPink();
		let cf = conds.AddAverageCondition(wasmModule.AverageType.AboveStdDev3);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_LightSkyBlue();

		sheet.Range.get("A15").NumberValue = 12;
		sheet.Range.get("B15").NumberValue = 15;
		sheet.Range.get("C15").NumberValue = 18;
		sheet.Range.get("A16").NumberValue = 14;
		sheet.Range.get("B16").NumberValue = 17;
		sheet.Range.get("C16").NumberValue = 20;
	}

	function AddTop10_1(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A17:C20"));
		sheet.Range.get("A17:C20").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A17:C20").Style.Color = wasmModule.Color.get_Gray();
		let cf = conds.AddTopBottomCondition(wasmModule.TopBottomType.Top, 10);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Yellow();

		sheet.Range.get("A17").NumberValue = 16;
		sheet.Range.get("B17").NumberValue = 21;
		sheet.Range.get("C17").NumberValue = 26;
		sheet.Range.get("A18").NumberValue = 18;
		sheet.Range.get("B18").NumberValue = 23;
		sheet.Range.get("C18").NumberValue = 28;
		sheet.Range.get("A19").NumberValue = 20;
		sheet.Range.get("B19").NumberValue = 25;
		sheet.Range.get("C19").NumberValue = 30;
		sheet.Range.get("A20").NumberValue = 22;
		sheet.Range.get("B20").NumberValue = 27;
		sheet.Range.get("C20").NumberValue = 32;
	}

	function AddTop10_2(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A21:C24"));
		sheet.Range.get("A21:C24").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A21:C24").Style.Color = wasmModule.Color.get_Green();
		let cf = conds.AddTopBottomCondition(wasmModule.TopBottomType.Bottom, 10);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Pink();

		sheet.Range.get("A21").NumberValue = 20;
		sheet.Range.get("B21").NumberValue = 25;
		sheet.Range.get("C21").NumberValue = 30;
		sheet.Range.get("A22").NumberValue = 22;
		sheet.Range.get("B22").NumberValue = 27;
		sheet.Range.get("C22").NumberValue = 32;
		sheet.Range.get("A23").NumberValue = 24;
		sheet.Range.get("B23").NumberValue = 29;
		sheet.Range.get("C23").NumberValue = 34;
		sheet.Range.get("A24").NumberValue = 24;
		sheet.Range.get("B24").NumberValue = 31;
		sheet.Range.get("C24").NumberValue = 36;
	}

	function AddTop10_3(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A25:C28"));
		sheet.Range.get("A25:C28").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A25:C28").Style.Color = wasmModule.Color.get_Orange();
		let cf = conds.AddTopBottomCondition(wasmModule.TopBottomType.TopPercent, 10);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Blue();

		sheet.Range.get("A25").NumberValue = 24;
		sheet.Range.get("B25").NumberValue = 29;
		sheet.Range.get("C25").NumberValue = 34;
		sheet.Range.get("A26").NumberValue = 25;
		sheet.Range.get("B26").NumberValue = 36;
		sheet.Range.get("C26").NumberValue = 32;
		sheet.Range.get("A27").NumberValue = 24;
		sheet.Range.get("B27").NumberValue = 28;
		sheet.Range.get("C27").NumberValue = 31;
		sheet.Range.get("A28").NumberValue = 34;
		sheet.Range.get("B28").NumberValue = 26;
		sheet.Range.get("C28").NumberValue = 32;
	}

	function AddTop10_4(sheet) {
		let conds = sheet.ConditionalFormats.Add();
		conds.AddRange(sheet.Range.get("A29:C32"));
		sheet.Range.get("A29:C32").Style.FillPattern = wasmModule.ExcelPatternType.Solid;
		sheet.Range.get("A29:C32").Style.Color = wasmModule.Color.get_Gold();
		let cf = conds.AddTopBottomCondition(wasmModule.TopBottomType.BottomPercent, 10);
		cf.FillPattern = wasmModule.ExcelPatternType.Solid;
		cf.BackColor = wasmModule.Color.get_Green();

		sheet.Range.get("A29").NumberValue = 22;
		sheet.Range.get("B29").NumberValue = 33;
		sheet.Range.get("C29").NumberValue = 38;
		sheet.Range.get("A30").NumberValue = 30;
		sheet.Range.get("B30").NumberValue = 35;
		sheet.Range.get("C30").NumberValue = 39;
		sheet.Range.get("A31").NumberValue = 32;
		sheet.Range.get("B31").NumberValue = 37;
		sheet.Range.get("C31").NumberValue = 43;
		sheet.Range.get("A32").NumberValue = 34;
		sheet.Range.get("B32").NumberValue = 28;
		sheet.Range.get("C32").NumberValue = 32;
	}

        //Create a workbook.
        const workbook = wasmModule.Workbook.Create();

        //Create a blank sheet
        workbook.CreateEmptySheets(1);
		
        //Get the first sheet
        let sheet = workbook.Worksheets.get(0);
		
        AddConditionalFormattingForNewSheet(sheet);

        //Save result file
        const outputFileName = 'VariousConditionalFormatting_out.xlsx';
        workbook.SaveToFile({fileName: outputFileName, version:wasmModule.ExcelVersion.Version2010});

        //Dispose
        workbook.Dispose();
		
        // Read the saved file and convert it to Bolb
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

        // Download the result file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };

    return {
      startProcessing,
      downloadName,
      downloadUrl
    };
  }
};
</script>
