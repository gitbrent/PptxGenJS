// Type definitions for pptxgenjs 3.7.0
// Project: https://gitbrent.github.io/PptxGenJS/
// Definitions by: Brent Ely <https://github.com/gitbrent/>
//                 Michael Beaumont <https://github.com/michaelbeaumont>
//                 Nicholas Tietz-Sokolsky <https://github.com/ntietz>
//                 David Adams <https://github.com/iota-pi>
//                 Stephen Cronin <https://github.com/cronin4392>
// TypeScript Version: 3.x

export as namespace PptxGenJS

export default PptxGenJS

declare class PptxGenJS {
	/**
	 * PptxGenJS Library Version
	 * @type {string}
	 */
	readonly version: string

	// Exposed prop types
	readonly presLayout: PptxGenJS.PresentationProps
	readonly AlignH: typeof PptxGenJS.AlignH
	readonly AlignV: typeof PptxGenJS.AlignV
	readonly ChartType: typeof PptxGenJS.ChartType
	readonly OutputType: typeof PptxGenJS.OutputType
	readonly SchemeColor: typeof PptxGenJS.SchemeColor
	readonly ShapeType: typeof PptxGenJS.ShapeType
	readonly PlaceholderType: typeof PptxGenJS.PLACEHOLDER_TYPES

	// Presentation Props

	/**
	 * Presentation layout name.
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10" x 7.5")
	 * - 'LAYOUT_16x9'  (10" x 5.625")
	 * - 'LAYOUT_16x10' (10" x 6.25")
	 * - 'LAYOUT_WIDE'  (13.33" x 7.5")
	 *
	 * Custom layouts:
	 * - Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 *
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	layout: string
	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	rtlMode: boolean

	// Presentation Metadata
	/**
	 * Author name
	 * @type {string}
	 */
	author: string
	/**
	 * Comapny name
	 * @type {string}
	 */
	company: string
	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PowerPoint will throw errors upon opening!)
	 */
	revision: string
	/**
	 * Presentation subject
	 * @type {string}
	 */
	subject: string
	/**
	 * Presentation name
	 * @type {string}
	 */
	title: string

	// Methods

	/**
	 * Export the current Presentation to stream
	 * @param {WriteBaseProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file stream
	 */
	stream(props?: PptxGenJS.WriteBaseProps): Promise<string | ArrayBuffer | Blob | Uint8Array>
	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file content in selected type
	 */
	write(props?: PptxGenJS.WriteProps): Promise<string | ArrayBuffer | Blob | Uint8Array>
	/**
	 * Export the current Presentation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
	 * @param {WriteFileProps} props output file properties
	 * @example pptx.writeFile({ fileName:'CustomerReport.pptx' }) // export presentation as "CustomerReport.pptx"
	 * @example pptx.writeFile({ fileName:'CustomerReport.pptx', compression:true }) // export presentation as "CustomerReport.pptx" compressed (can save up to 30%)
	 * @returns {Promise<string>} the presentation name
	 */
	writeFile(props?: PptxGenJS.WriteFileProps): Promise<string>
	/**
	 * Add a new Section to Presentation
	 * @param {SectionProps} props section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(props: PptxGenJS.SectionProps): void
	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} props slide options
	 * @returns {Slide} the new Slide
	 */
	addSlide(props?: PptxGenJS.AddSlideProps): PptxGenJS.Slide
	/**
	 * Add a new Slide to Presentation
	 * @param {string} masterName master slide name
	 * @returns {Slide} the new Slide
	 * @deprecated use `addSlide(IAddSlideOptions)`
	 */
	addSlide(masterName?: string): PptxGenJS.Slide
	/**
	 * Create a custom Slide Layout in any size
	 * @param {PresLayout} layout an object with user-defined w/h
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: PptxGenJS.PresLayout): void
	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props layout definition
	 */
	defineSlideMaster(props: PptxGenJS.SlideMasterProps): void
	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId table HTML element ID
	 * @param {TableToSlidesProps} props generation options
	 */
	tableToSlides(eleId: string, props?: PptxGenJS.TableToSlidesProps): void
}

declare namespace PptxGenJS {
	// Exported enums for module apps
	// @example: pptxgen.ShapeType.rect
	export enum AlignH {
		'left' = 'left',
		'center' = 'center',
		'right' = 'right',
		'justify' = 'justify',
	}
	export enum AlignV {
		'top' = 'top',
		'middle' = 'middle',
		'bottom' = 'bottom',
	}
	export enum ChartType {
		'area' = 'area',
		'bar' = 'bar',
		'bar3d' = 'bar3D',
		'bubble' = 'bubble',
		'doughnut' = 'doughnut',
		'line' = 'line',
		'pie' = 'pie',
		'radar' = 'radar',
		'scatter' = 'scatter',
	}
	export enum OutputType {
		'arraybuffer' = 'arraybuffer',
		'base64' = 'base64',
		'binarystring' = 'binarystring',
		'blob' = 'blob',
		'nodebuffer' = 'nodebuffer',
		'uint8array' = 'uint8array',
	}
	export enum SchemeColor {
		'text1' = 'tx1',
		'text2' = 'tx2',
		'background1' = 'bg1',
		'background2' = 'bg2',
		'accent1' = 'accent1',
		'accent2' = 'accent2',
		'accent3' = 'accent3',
		'accent4' = 'accent4',
		'accent5' = 'accent5',
		'accent6' = 'accent6',
	}
	export enum ShapeType {
		'accentBorderCallout1' = 'accentBorderCallout1',
		'accentBorderCallout2' = 'accentBorderCallout2',
		'accentBorderCallout3' = 'accentBorderCallout3',
		'accentCallout1' = 'accentCallout1',
		'accentCallout2' = 'accentCallout2',
		'accentCallout3' = 'accentCallout3',
		'actionButtonBackPrevious' = 'actionButtonBackPrevious',
		'actionButtonBeginning' = 'actionButtonBeginning',
		'actionButtonBlank' = 'actionButtonBlank',
		'actionButtonDocument' = 'actionButtonDocument',
		'actionButtonEnd' = 'actionButtonEnd',
		'actionButtonForwardNext' = 'actionButtonForwardNext',
		'actionButtonHelp' = 'actionButtonHelp',
		'actionButtonHome' = 'actionButtonHome',
		'actionButtonInformation' = 'actionButtonInformation',
		'actionButtonMovie' = 'actionButtonMovie',
		'actionButtonReturn' = 'actionButtonReturn',
		'actionButtonSound' = 'actionButtonSound',
		'arc' = 'arc',
		'bentArrow' = 'bentArrow',
		'bentUpArrow' = 'bentUpArrow',
		'bevel' = 'bevel',
		'blockArc' = 'blockArc',
		'borderCallout1' = 'borderCallout1',
		'borderCallout2' = 'borderCallout2',
		'borderCallout3' = 'borderCallout3',
		'bracePair' = 'bracePair',
		'bracketPair' = 'bracketPair',
		'callout1' = 'callout1',
		'callout2' = 'callout2',
		'callout3' = 'callout3',
		'can' = 'can',
		'chartPlus' = 'chartPlus',
		'chartStar' = 'chartStar',
		'chartX' = 'chartX',
		'chevron' = 'chevron',
		'chord' = 'chord',
		'circularArrow' = 'circularArrow',
		'cloud' = 'cloud',
		'cloudCallout' = 'cloudCallout',
		'corner' = 'corner',
		'cornerTabs' = 'cornerTabs',
		'cube' = 'cube',
		'curvedDownArrow' = 'curvedDownArrow',
		'curvedLeftArrow' = 'curvedLeftArrow',
		'curvedRightArrow' = 'curvedRightArrow',
		'curvedUpArrow' = 'curvedUpArrow',
		'decagon' = 'decagon',
		'diagStripe' = 'diagStripe',
		'diamond' = 'diamond',
		'dodecagon' = 'dodecagon',
		'donut' = 'donut',
		'doubleWave' = 'doubleWave',
		'downArrow' = 'downArrow',
		'downArrowCallout' = 'downArrowCallout',
		'ellipse' = 'ellipse',
		'ellipseRibbon' = 'ellipseRibbon',
		'ellipseRibbon2' = 'ellipseRibbon2',
		'flowChartAlternateProcess' = 'flowChartAlternateProcess',
		'flowChartCollate' = 'flowChartCollate',
		'flowChartConnector' = 'flowChartConnector',
		'flowChartDecision' = 'flowChartDecision',
		'flowChartDelay' = 'flowChartDelay',
		'flowChartDisplay' = 'flowChartDisplay',
		'flowChartDocument' = 'flowChartDocument',
		'flowChartExtract' = 'flowChartExtract',
		'flowChartInputOutput' = 'flowChartInputOutput',
		'flowChartInternalStorage' = 'flowChartInternalStorage',
		'flowChartMagneticDisk' = 'flowChartMagneticDisk',
		'flowChartMagneticDrum' = 'flowChartMagneticDrum',
		'flowChartMagneticTape' = 'flowChartMagneticTape',
		'flowChartManualInput' = 'flowChartManualInput',
		'flowChartManualOperation' = 'flowChartManualOperation',
		'flowChartMerge' = 'flowChartMerge',
		'flowChartMultidocument' = 'flowChartMultidocument',
		'flowChartOfflineStorage' = 'flowChartOfflineStorage',
		'flowChartOffpageConnector' = 'flowChartOffpageConnector',
		'flowChartOnlineStorage' = 'flowChartOnlineStorage',
		'flowChartOr' = 'flowChartOr',
		'flowChartPredefinedProcess' = 'flowChartPredefinedProcess',
		'flowChartPreparation' = 'flowChartPreparation',
		'flowChartProcess' = 'flowChartProcess',
		'flowChartPunchedCard' = 'flowChartPunchedCard',
		'flowChartPunchedTape' = 'flowChartPunchedTape',
		'flowChartSort' = 'flowChartSort',
		'flowChartSummingJunction' = 'flowChartSummingJunction',
		'flowChartTerminator' = 'flowChartTerminator',
		'folderCorner' = 'folderCorner',
		'frame' = 'frame',
		'funnel' = 'funnel',
		'gear6' = 'gear6',
		'gear9' = 'gear9',
		'halfFrame' = 'halfFrame',
		'heart' = 'heart',
		'heptagon' = 'heptagon',
		'hexagon' = 'hexagon',
		'homePlate' = 'homePlate',
		'horizontalScroll' = 'horizontalScroll',
		'irregularSeal1' = 'irregularSeal1',
		'irregularSeal2' = 'irregularSeal2',
		'leftArrow' = 'leftArrow',
		'leftArrowCallout' = 'leftArrowCallout',
		'leftBrace' = 'leftBrace',
		'leftBracket' = 'leftBracket',
		'leftCircularArrow' = 'leftCircularArrow',
		'leftRightArrow' = 'leftRightArrow',
		'leftRightArrowCallout' = 'leftRightArrowCallout',
		'leftRightCircularArrow' = 'leftRightCircularArrow',
		'leftRightRibbon' = 'leftRightRibbon',
		'leftRightUpArrow' = 'leftRightUpArrow',
		'leftUpArrow' = 'leftUpArrow',
		'lightningBolt' = 'lightningBolt',
		'line' = 'line',
		'lineInv' = 'lineInv',
		'mathDivide' = 'mathDivide',
		'mathEqual' = 'mathEqual',
		'mathMinus' = 'mathMinus',
		'mathMultiply' = 'mathMultiply',
		'mathNotEqual' = 'mathNotEqual',
		'mathPlus' = 'mathPlus',
		'moon' = 'moon',
		'nonIsoscelesTrapezoid' = 'nonIsoscelesTrapezoid',
		'noSmoking' = 'noSmoking',
		'notchedRightArrow' = 'notchedRightArrow',
		'octagon' = 'octagon',
		'parallelogram' = 'parallelogram',
		'pentagon' = 'pentagon',
		'pie' = 'pie',
		'pieWedge' = 'pieWedge',
		'plaque' = 'plaque',
		'plaqueTabs' = 'plaqueTabs',
		'plus' = 'plus',
		'quadArrow' = 'quadArrow',
		'quadArrowCallout' = 'quadArrowCallout',
		'rect' = 'rect',
		'ribbon' = 'ribbon',
		'ribbon2' = 'ribbon2',
		'rightArrow' = 'rightArrow',
		'rightArrowCallout' = 'rightArrowCallout',
		'rightBrace' = 'rightBrace',
		'rightBracket' = 'rightBracket',
		'round1Rect' = 'round1Rect',
		'round2DiagRect' = 'round2DiagRect',
		'round2SameRect' = 'round2SameRect',
		'roundRect' = 'roundRect',
		'rtTriangle' = 'rtTriangle',
		'smileyFace' = 'smileyFace',
		'snip1Rect' = 'snip1Rect',
		'snip2DiagRect' = 'snip2DiagRect',
		'snip2SameRect' = 'snip2SameRect',
		'snipRoundRect' = 'snipRoundRect',
		'squareTabs' = 'squareTabs',
		'star10' = 'star10',
		'star12' = 'star12',
		'star16' = 'star16',
		'star24' = 'star24',
		'star32' = 'star32',
		'star4' = 'star4',
		'star5' = 'star5',
		'star6' = 'star6',
		'star7' = 'star7',
		'star8' = 'star8',
		'stripedRightArrow' = 'stripedRightArrow',
		'sun' = 'sun',
		'swooshArrow' = 'swooshArrow',
		'teardrop' = 'teardrop',
		'trapezoid' = 'trapezoid',
		'triangle' = 'triangle',
		'upArrow' = 'upArrow',
		'upArrowCallout' = 'upArrowCallout',
		'upDownArrow' = 'upDownArrow',
		'upDownArrowCallout' = 'upDownArrowCallout',
		'uturnArrow' = 'uturnArrow',
		'verticalScroll' = 'verticalScroll',
		'wave' = 'wave',
		'wedgeEllipseCallout' = 'wedgeEllipseCallout',
		'wedgeRectCallout' = 'wedgeRectCallout',
		'wedgeRoundRectCallout' = 'wedgeRoundRectCallout',
	}
	// used by charts, shape, text
	export interface BorderOptions {
		/**
		 * Border type
		 */
		type?: 'none' | 'dash' | 'solid'
		/**
		 * Border color (hex)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * Border size (points)
		 */
		pt?: number
	}
	// These are used by browser/script clients and have been named like this since v0.1.
	// Desc: charts and shapes for `pptxgen.charts.` `pptxgen.shapes.`
	// Note: "charts" and "shapes" are manually created by cloning
	export enum charts {
		'AREA' = 'area',
		'BAR' = 'bar',
		'BAR3D' = 'bar3D',
		'BUBBLE' = 'bubble',
		'DOUGHNUT' = 'doughnut',
		'LINE' = 'line',
		'PIE' = 'pie',
		'RADAR' = 'radar',
		'SCATTER' = 'scatter',
	}
	export enum shapes {
		ACTION_BUTTON_BACK_OR_PREVIOUS = 'actionButtonBackPrevious',
		ACTION_BUTTON_BEGINNING = 'actionButtonBeginning',
		ACTION_BUTTON_CUSTOM = 'actionButtonBlank',
		ACTION_BUTTON_DOCUMENT = 'actionButtonDocument',
		ACTION_BUTTON_END = 'actionButtonEnd',
		ACTION_BUTTON_FORWARD_OR_NEXT = 'actionButtonForwardNext',
		ACTION_BUTTON_HELP = 'actionButtonHelp',
		ACTION_BUTTON_HOME = 'actionButtonHome',
		ACTION_BUTTON_INFORMATION = 'actionButtonInformation',
		ACTION_BUTTON_MOVIE = 'actionButtonMovie',
		ACTION_BUTTON_RETURN = 'actionButtonReturn',
		ACTION_BUTTON_SOUND = 'actionButtonSound',
		ARC = 'arc',
		BALLOON = 'wedgeRoundRectCallout',
		BENT_ARROW = 'bentArrow',
		BENT_UP_ARROW = 'bentUpArrow',
		BEVEL = 'bevel',
		BLOCK_ARC = 'blockArc',
		CAN = 'can',
		CHART_PLUS = 'chartPlus',
		CHART_STAR = 'chartStar',
		CHART_X = 'chartX',
		CHEVRON = 'chevron',
		CHORD = 'chord',
		CIRCULAR_ARROW = 'circularArrow',
		CLOUD = 'cloud',
		CLOUD_CALLOUT = 'cloudCallout',
		CORNER = 'corner',
		CORNER_TABS = 'cornerTabs',
		CROSS = 'plus',
		CUBE = 'cube',
		CURVED_DOWN_ARROW = 'curvedDownArrow',
		CURVED_DOWN_RIBBON = 'ellipseRibbon',
		CURVED_LEFT_ARROW = 'curvedLeftArrow',
		CURVED_RIGHT_ARROW = 'curvedRightArrow',
		CURVED_UP_ARROW = 'curvedUpArrow',
		CURVED_UP_RIBBON = 'ellipseRibbon2',
		DECAGON = 'decagon',
		DIAGONAL_STRIPE = 'diagStripe',
		DIAMOND = 'diamond',
		DODECAGON = 'dodecagon',
		DONUT = 'donut',
		DOUBLE_BRACE = 'bracePair',
		DOUBLE_BRACKET = 'bracketPair',
		DOUBLE_WAVE = 'doubleWave',
		DOWN_ARROW = 'downArrow',
		DOWN_ARROW_CALLOUT = 'downArrowCallout',
		DOWN_RIBBON = 'ribbon',
		EXPLOSION1 = 'irregularSeal1',
		EXPLOSION2 = 'irregularSeal2',
		FLOWCHART_ALTERNATE_PROCESS = 'flowChartAlternateProcess',
		FLOWCHART_CARD = 'flowChartPunchedCard',
		FLOWCHART_COLLATE = 'flowChartCollate',
		FLOWCHART_CONNECTOR = 'flowChartConnector',
		FLOWCHART_DATA = 'flowChartInputOutput',
		FLOWCHART_DECISION = 'flowChartDecision',
		FLOWCHART_DELAY = 'flowChartDelay',
		FLOWCHART_DIRECT_ACCESS_STORAGE = 'flowChartMagneticDrum',
		FLOWCHART_DISPLAY = 'flowChartDisplay',
		FLOWCHART_DOCUMENT = 'flowChartDocument',
		FLOWCHART_EXTRACT = 'flowChartExtract',
		FLOWCHART_INTERNAL_STORAGE = 'flowChartInternalStorage',
		FLOWCHART_MAGNETIC_DISK = 'flowChartMagneticDisk',
		FLOWCHART_MANUAL_INPUT = 'flowChartManualInput',
		FLOWCHART_MANUAL_OPERATION = 'flowChartManualOperation',
		FLOWCHART_MERGE = 'flowChartMerge',
		FLOWCHART_MULTIDOCUMENT = 'flowChartMultidocument',
		FLOWCHART_OFFLINE_STORAGE = 'flowChartOfflineStorage',
		FLOWCHART_OFFPAGE_CONNECTOR = 'flowChartOffpageConnector',
		FLOWCHART_OR = 'flowChartOr',
		FLOWCHART_PREDEFINED_PROCESS = 'flowChartPredefinedProcess',
		FLOWCHART_PREPARATION = 'flowChartPreparation',
		FLOWCHART_PROCESS = 'flowChartProcess',
		FLOWCHART_PUNCHED_TAPE = 'flowChartPunchedTape',
		FLOWCHART_SEQUENTIAL_ACCESS_STORAGE = 'flowChartMagneticTape',
		FLOWCHART_SORT = 'flowChartSort',
		FLOWCHART_STORED_DATA = 'flowChartOnlineStorage',
		FLOWCHART_SUMMING_JUNCTION = 'flowChartSummingJunction',
		FLOWCHART_TERMINATOR = 'flowChartTerminator',
		FOLDED_CORNER = 'folderCorner',
		FRAME = 'frame',
		FUNNEL = 'funnel',
		GEAR_6 = 'gear6',
		GEAR_9 = 'gear9',
		HALF_FRAME = 'halfFrame',
		HEART = 'heart',
		HEPTAGON = 'heptagon',
		HEXAGON = 'hexagon',
		HORIZONTAL_SCROLL = 'horizontalScroll',
		ISOSCELES_TRIANGLE = 'triangle',
		LEFT_ARROW = 'leftArrow',
		LEFT_ARROW_CALLOUT = 'leftArrowCallout',
		LEFT_BRACE = 'leftBrace',
		LEFT_BRACKET = 'leftBracket',
		LEFT_CIRCULAR_ARROW = 'leftCircularArrow',
		LEFT_RIGHT_ARROW = 'leftRightArrow',
		LEFT_RIGHT_ARROW_CALLOUT = 'leftRightArrowCallout',
		LEFT_RIGHT_CIRCULAR_ARROW = 'leftRightCircularArrow',
		LEFT_RIGHT_RIBBON = 'leftRightRibbon',
		LEFT_RIGHT_UP_ARROW = 'leftRightUpArrow',
		LEFT_UP_ARROW = 'leftUpArrow',
		LIGHTNING_BOLT = 'lightningBolt',
		LINE_CALLOUT_1 = 'borderCallout1',
		LINE_CALLOUT_1_ACCENT_BAR = 'accentCallout1',
		LINE_CALLOUT_1_BORDER_AND_ACCENT_BAR = 'accentBorderCallout1',
		LINE_CALLOUT_1_NO_BORDER = 'callout1',
		LINE_CALLOUT_2 = 'borderCallout2',
		LINE_CALLOUT_2_ACCENT_BAR = 'accentCallout2',
		LINE_CALLOUT_2_BORDER_AND_ACCENT_BAR = 'accentBorderCallout2',
		LINE_CALLOUT_2_NO_BORDER = 'callout2',
		LINE_CALLOUT_3 = 'borderCallout3',
		LINE_CALLOUT_3_ACCENT_BAR = 'accentCallout3',
		LINE_CALLOUT_3_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3',
		LINE_CALLOUT_3_NO_BORDER = 'callout3',
		LINE_CALLOUT_4 = 'borderCallout3',
		LINE_CALLOUT_4_ACCENT_BAR = 'accentCallout3',
		LINE_CALLOUT_4_BORDER_AND_ACCENT_BAR = 'accentBorderCallout3',
		LINE_CALLOUT_4_NO_BORDER = 'callout3',
		LINE = 'line',
		LINE_INVERSE = 'lineInv',
		MATH_DIVIDE = 'mathDivide',
		MATH_EQUAL = 'mathEqual',
		MATH_MINUS = 'mathMinus',
		MATH_MULTIPLY = 'mathMultiply',
		MATH_NOT_EQUAL = 'mathNotEqual',
		MATH_PLUS = 'mathPlus',
		MOON = 'moon',
		NON_ISOSCELES_TRAPEZOID = 'nonIsoscelesTrapezoid',
		NOTCHED_RIGHT_ARROW = 'notchedRightArrow',
		NO_SYMBOL = 'noSmoking',
		OCTAGON = 'octagon',
		OVAL = 'ellipse',
		OVAL_CALLOUT = 'wedgeEllipseCallout',
		PARALLELOGRAM = 'parallelogram',
		PENTAGON = 'homePlate',
		PIE = 'pie',
		PIE_WEDGE = 'pieWedge',
		PLAQUE = 'plaque',
		PLAQUE_TABS = 'plaqueTabs',
		QUAD_ARROW = 'quadArrow',
		QUAD_ARROW_CALLOUT = 'quadArrowCallout',
		RECTANGLE = 'rect',
		RECTANGULAR_CALLOUT = 'wedgeRectCallout',
		REGULAR_PENTAGON = 'pentagon',
		RIGHT_ARROW = 'rightArrow',
		RIGHT_ARROW_CALLOUT = 'rightArrowCallout',
		RIGHT_BRACE = 'rightBrace',
		RIGHT_BRACKET = 'rightBracket',
		RIGHT_TRIANGLE = 'rtTriangle',
		ROUNDED_RECTANGLE = 'roundRect',
		ROUNDED_RECTANGULAR_CALLOUT = 'wedgeRoundRectCallout',
		ROUND_1_RECTANGLE = 'round1Rect',
		ROUND_2_DIAG_RECTANGLE = 'round2DiagRect',
		ROUND_2_SAME_RECTANGLE = 'round2SameRect',
		SMILEY_FACE = 'smileyFace',
		SNIP_1_RECTANGLE = 'snip1Rect',
		SNIP_2_DIAG_RECTANGLE = 'snip2DiagRect',
		SNIP_2_SAME_RECTANGLE = 'snip2SameRect',
		SNIP_ROUND_RECTANGLE = 'snipRoundRect',
		SQUARE_TABS = 'squareTabs',
		STAR_10_POINT = 'star10',
		STAR_12_POINT = 'star12',
		STAR_16_POINT = 'star16',
		STAR_24_POINT = 'star24',
		STAR_32_POINT = 'star32',
		STAR_4_POINT = 'star4',
		STAR_5_POINT = 'star5',
		STAR_6_POINT = 'star6',
		STAR_7_POINT = 'star7',
		STAR_8_POINT = 'star8',
		STRIPED_RIGHT_ARROW = 'stripedRightArrow',
		SUN = 'sun',
		SWOOSH_ARROW = 'swooshArrow',
		TEAR = 'teardrop',
		TRAPEZOID = 'trapezoid',
		UP_ARROW = 'upArrow',
		UP_ARROW_CALLOUT = 'upArrowCallout',
		UP_DOWN_ARROW = 'upDownArrow',
		UP_DOWN_ARROW_CALLOUT = 'upDownArrowCallout',
		UP_RIBBON = 'ribbon2',
		U_TURN_ARROW = 'uturnArrow',
		VERTICAL_SCROLL = 'verticalScroll',
		WAVE = 'wave',
	}

	// @source `core-enums.ts`
	export type JSZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	export type WRITE_OUTPUT_TYPE = JSZIP_OUTPUT_TYPE | 'STREAM'
	export enum CHART_TYPE {
		'AREA' = 'area',
		'BAR' = 'bar',
		'BAR3D' = 'bar3D',
		'BUBBLE' = 'bubble',
		'DOUGHNUT' = 'doughnut',
		'LINE' = 'line',
		'PIE' = 'pie',
		'RADAR' = 'radar',
		'SCATTER' = 'scatter',
	}
	export enum SCHEME_COLOR_NAMES {
		'TEXT1' = 'tx1',
		'TEXT2' = 'tx2',
		'BACKGROUND1' = 'bg1',
		'BACKGROUND2' = 'bg2',
		'ACCENT1' = 'accent1',
		'ACCENT2' = 'accent2',
		'ACCENT3' = 'accent3',
		'ACCENT4' = 'accent4',
		'ACCENT5' = 'accent5',
		'ACCENT6' = 'accent6',
	}

	// @source `core-interfaces.d.ts` (via import)
	// @code `import { CHART_NAME, PLACEHOLDER_TYPES, SHAPE_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN } from './core-enums'`
	export type CHART_NAME = 'area' | 'bar' | 'bar3D' | 'bubble' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'
	export enum PLACEHOLDER_TYPES {
		'title' = 'title',
		'body' = 'body',
		'image' = 'pic',
		'chart' = 'chart',
		'table' = 'tbl',
		'media' = 'media',
	}
	export type PLACEHOLDER_TYPE = 'title' | 'body' | 'pic' | 'chart' | 'tbl' | 'media'

	export type SHAPE_NAME =
		| 'accentBorderCallout1'
		| 'accentBorderCallout2'
		| 'accentBorderCallout3'
		| 'accentCallout1'
		| 'accentCallout2'
		| 'accentCallout3'
		| 'actionButtonBackPrevious'
		| 'actionButtonBeginning'
		| 'actionButtonBlank'
		| 'actionButtonDocument'
		| 'actionButtonEnd'
		| 'actionButtonForwardNext'
		| 'actionButtonHelp'
		| 'actionButtonHome'
		| 'actionButtonInformation'
		| 'actionButtonMovie'
		| 'actionButtonReturn'
		| 'actionButtonSound'
		| 'arc'
		| 'bentArrow'
		| 'bentUpArrow'
		| 'bevel'
		| 'blockArc'
		| 'borderCallout1'
		| 'borderCallout2'
		| 'borderCallout3'
		| 'bracePair'
		| 'bracketPair'
		| 'callout1'
		| 'callout2'
		| 'callout3'
		| 'can'
		| 'chartPlus'
		| 'chartStar'
		| 'chartX'
		| 'chevron'
		| 'chord'
		| 'circularArrow'
		| 'cloud'
		| 'cloudCallout'
		| 'corner'
		| 'cornerTabs'
		| 'cube'
		| 'curvedDownArrow'
		| 'curvedLeftArrow'
		| 'curvedRightArrow'
		| 'curvedUpArrow'
		| 'decagon'
		| 'diagStripe'
		| 'diamond'
		| 'dodecagon'
		| 'donut'
		| 'doubleWave'
		| 'downArrow'
		| 'downArrowCallout'
		| 'ellipse'
		| 'ellipseRibbon'
		| 'ellipseRibbon2'
		| 'flowChartAlternateProcess'
		| 'flowChartCollate'
		| 'flowChartConnector'
		| 'flowChartDecision'
		| 'flowChartDelay'
		| 'flowChartDisplay'
		| 'flowChartDocument'
		| 'flowChartExtract'
		| 'flowChartInputOutput'
		| 'flowChartInternalStorage'
		| 'flowChartMagneticDisk'
		| 'flowChartMagneticDrum'
		| 'flowChartMagneticTape'
		| 'flowChartManualInput'
		| 'flowChartManualOperation'
		| 'flowChartMerge'
		| 'flowChartMultidocument'
		| 'flowChartOfflineStorage'
		| 'flowChartOffpageConnector'
		| 'flowChartOnlineStorage'
		| 'flowChartOr'
		| 'flowChartPredefinedProcess'
		| 'flowChartPreparation'
		| 'flowChartProcess'
		| 'flowChartPunchedCard'
		| 'flowChartPunchedTape'
		| 'flowChartSort'
		| 'flowChartSummingJunction'
		| 'flowChartTerminator'
		| 'folderCorner'
		| 'frame'
		| 'funnel'
		| 'gear6'
		| 'gear9'
		| 'halfFrame'
		| 'heart'
		| 'heptagon'
		| 'hexagon'
		| 'homePlate'
		| 'horizontalScroll'
		| 'irregularSeal1'
		| 'irregularSeal2'
		| 'leftArrow'
		| 'leftArrowCallout'
		| 'leftBrace'
		| 'leftBracket'
		| 'leftCircularArrow'
		| 'leftRightArrow'
		| 'leftRightArrowCallout'
		| 'leftRightCircularArrow'
		| 'leftRightRibbon'
		| 'leftRightUpArrow'
		| 'leftUpArrow'
		| 'lightningBolt'
		| 'line'
		| 'lineInv'
		| 'mathDivide'
		| 'mathEqual'
		| 'mathMinus'
		| 'mathMultiply'
		| 'mathNotEqual'
		| 'mathPlus'
		| 'moon'
		| 'noSmoking'
		| 'nonIsoscelesTrapezoid'
		| 'notchedRightArrow'
		| 'octagon'
		| 'parallelogram'
		| 'pentagon'
		| 'pie'
		| 'pieWedge'
		| 'plaque'
		| 'plaqueTabs'
		| 'plus'
		| 'quadArrow'
		| 'quadArrowCallout'
		| 'rect'
		| 'ribbon'
		| 'ribbon2'
		| 'rightArrow'
		| 'rightArrowCallout'
		| 'rightBrace'
		| 'rightBracket'
		| 'round1Rect'
		| 'round2DiagRect'
		| 'round2SameRect'
		| 'roundRect'
		| 'rtTriangle'
		| 'smileyFace'
		| 'snip1Rect'
		| 'snip2DiagRect'
		| 'snip2SameRect'
		| 'snipRoundRect'
		| 'squareTabs'
		| 'star10'
		| 'star12'
		| 'star16'
		| 'star24'
		| 'star32'
		| 'star4'
		| 'star5'
		| 'star6'
		| 'star7'
		| 'star8'
		| 'stripedRightArrow'
		| 'sun'
		| 'swooshArrow'
		| 'teardrop'
		| 'trapezoid'
		| 'triangle'
		| 'upArrow'
		| 'upArrowCallout'
		| 'upDownArrow'
		| 'upDownArrowCallout'
		| 'uturnArrow'
		| 'verticalScroll'
		| 'wave'
		| 'wedgeEllipseCallout'
		| 'wedgeRectCallout'
		| 'wedgeRoundRectCallout'

	export enum SLIDE_OBJECT_TYPES {
		'chart' = 'chart',
		'hyperlink' = 'hyperlink',
		'image' = 'image',
		'media' = 'media',
		'online' = 'online',
		'placeholder' = 'placeholder',
		'table' = 'table',
		'tablecell' = 'tablecell',
		'text' = 'text',
		'notes' = 'notes',
	}
	export enum TEXT_HALIGN {
		'left' = 'left',
		'center' = 'center',
		'right' = 'right',
		'justify' = 'justify',
	}
	export enum TEXT_VALIGN {
		'b' = 'b',
		'ctr' = 'ctr',
		't' = 't',
	}

	// @source `core-interfaces.d.ts` (direct)
	// Core Types
	// ==========

	/**
	 * Coordinate number - either:
	 * - Inches
	 * - Percentage
	 *
	 * @example 10.25 // coordinate in inches
	 * @example '75%' // coordinate as percentage of slide size
	 */
	export type Coord = number | string
	export type PositionProps = {
		/**
		 * Horizontal position
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		x?: Coord
		/**
		 * Vertical position
		 * - inches or percentage
		 * @example 10.25 // position in inches
		 * @example '75%' // position as percentage of slide size
		 */
		y?: Coord
		/**
		 * Height
		 * - inches or percentage
		 * @example 10.25 // height in inches
		 * @example '75%' // height as percentage of slide size
		 */
		h?: Coord
		/**
		 * Width
		 * - inches or percentage
		 * @example 10.25 // width in inches
		 * @example '75%' // width as percentage of slide size
		 */
		w?: Coord
	}
	/**
	 * Either `data` or `path` is required
	 */
	export type DataOrPathProps = {
		/**
		 * URL or relative path
		 *
		 * @example 'https://onedrives.com/myimg.png` // retrieve image via URL
		 * @example '/home/gitbrent/images/myimg.png` // retrieve image via local path
		 */
		path?: string
		/**
		 * base64-encoded string
		 * - Useful for avoiding potential path/server issues
		 *
		 * @example 'image/png;base64,iVtDafDrBF[...]=' // pre-encoded image in base-64
		 */
		data?: string
	}
	export interface BackgroundProps extends DataOrPathProps, ShapeFillProps {
		/**
		 * Color (hex format)
		 * @deprecated v3.6.0 - use `ShapeFillProps` instead
		 */
		fill?: HexColor
	}
	/**
	 * Color in Hex format
	 * @example 'FF3399'
	 */
	export type HexColor = string
	export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
	export interface ModifiedThemeColor {
		baseColor: HexColor | ThemeColor

		alpha?: number
		alphaMod?: number
		alphaOff?: number
		blue?: number
		blueMod?: number
		blueOff?: number
		green?: number
		greenMod?: number
		greenOff?: number
		red?: number
		redMod?: number
		redOff?: number
		hue?: number
		hueMod?: number
		hueOff?: number
		lum?: number
		lumMod?: number
		lumOff?: number
		sat?: number
		satMod?: number
		satOff?: number
		shade?: number
		tint?: number

		comp?: boolean
		gray?: boolean
		inv?: boolean
		gamma?: boolean
	}
	export type Color = HexColor | ThemeColor | ModifiedThemeColor
	export type Margin = number | [number, number, number, number]
	export type HAlign = 'left' | 'center' | 'right' | 'justify'
	export type VAlign = 'top' | 'middle' | 'bottom'

	// used by charts, shape, text
	export interface BorderProps {
		/**
		 * Border type
		 * @default solid
		 */
		type?: 'none' | 'dash' | 'solid'
		/**
		 * Border color (hex)
		 * @example 'FF3399'
		 * @default '666666'
		 */
		color?: HexColor

		// TODO: add `width` - deprecate `pt`
		/**
		 * Border size (points)
		 * @default 1
		 */
		pt?: number
	}
	// used by: image, object, text,
	export interface HyperlinkProps {
		/**
		 * Slide number to link to
		 */
		slide?: number
		/**
		 * Url to link to
		 */
		url?: string
		/**
		 * Hyperlink Tooltip
		 */
		tooltip?: string
	}
	export interface PlaceholderProps {
		name: string
		type: PLACEHOLDER_TYPE
		x: Coord
		y: Coord
		w: Coord
		h: Coord
	}
	// used by: chart, text
	export interface ShadowProps {
		/**
		 * shadow type
		 * @default 'none'
		 */
		type: 'outer' | 'inner' | 'none'
		/**
		 * opacity (0.0 - 1.0)
		 * @example 0.5 // 50% opaque
		 */
		opacity?: number // TODO: "Transparency (0-100%)" in PPT // TODO: deprecate and add `transparency`
		/**
		 * blur (points)
		 * - range: 0-100
		 * @default 0
		 */
		blur?: number
		/**
		 * angle (degrees)
		 * - range: 0-359
		 * @default 0
		 */
		angle?: number
		/**
		 * shadow offset (points)
		 * - range: 0-200
		 * @default 0
		 */
		offset?: number // TODO: "Distance" in PPT
		/**
		 * shadow color (hex format)
		 * @example 'FF3399'
		 */
		color?: HexColor
	}
	// used by: shape, table, text
	export interface ShapeFillProps {
		/**
		 * Fill color
		 * - `HexColor` or `ThemeColor`
		 * @example 'FF0000' // red
		 * @example 'pptx.SchemeColor.text1' // Text1 Theme Color
		 */
		color?: Color
		/**
		 * Transparency (percent)
		 * - range: 0-100
		 * @default 0
		 */
		transparency?: number
		/**
		 * Fill type
		 * @default 'solid'
		 */
		type?: 'none' | 'solid'

		/**
		 * Transparency (percent)
		 * @deprecated v3.3.0 - use `transparency`
		 */
		alpha?: number
	}
	export interface ShapeLineProps extends ShapeFillProps {
		/**
		 * Line width (pt)
		 * @default 1
		 */
		width?: number
		/**
		 * Dash type
		 * @default 'solid'
		 */
		dashType?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * Begin arrow type
		 * @since v3.3.0
		 */
		beginArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * End arrow type
		 * @since v3.3.0
		 */
		endArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		// FUTURE: beginArrowSize (1-9)
		// FUTURE: endArrowSize (1-9)

		/**
		 * Dash type
		 * @deprecated v3.3.0 - use `dashType`
		 */
		lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * @deprecated v3.3.0 - use `beginArrowType`
		 */
		lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @deprecated v3.3.0 - use `endArrowType`
		 */
		lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * Line width (pt)
		 * @deprecated v3.3.0 - use `width`
		 */
		pt?: number
		/**
		 * Line size (pt)
		 * @deprecated v3.3.0 - use `width`
		 */
		size?: number
	}
	// used by: chart, slide, table, text
	export interface TextBaseProps {
		/**
		 * Horizontal alignment
		 * @default 'left'
		 */
		align?: HAlign
		/**
		 * Bold style
		 * @default false
		 */
		bold?: boolean
		/**
		 * Add a line-break
		 * @default false
		 */
		breakLine?: boolean
		/**
		 * Add a soft line-break (shift+enter) before line text content
		 * @default false
		 * @since v3.5.0
		 */
		softBreakBefore?: boolean
		/**
		 * Add standard or custom bullet
		 * - use `true` for standard bullet
		 * - pass object options for custom bullet
		 * @default false
		 */
		bullet?:
			| boolean
			| {
					/**
					 * Bullet type
					 * @default bullet
					 */
					type?: 'bullet' | 'number'
					/**
					 * Bullet character code (unicode)
					 * @since v3.3.0
					 * @example '25BA' // 'BLACK RIGHT-POINTING POINTER' (U+25BA)
					 */
					characterCode?: string
					/**
					 * Indentation (space between bullet and text) (points)
					 * @since v3.3.0
					 * @default 27 // DEF_BULLET_MARGIN
					 * @example 10 // Indents text 10 points from bullet
					 */
					indent?: number
					/**
					 * Number type
					 * @since v3.3.0
					 * @example 'romanLcParenR' // roman numerals lower-case with paranthesis right
					 */
					numberType?:
						| 'alphaLcParenBoth'
						| 'alphaLcParenR'
						| 'alphaLcPeriod'
						| 'alphaUcParenBoth'
						| 'alphaUcParenR'
						| 'alphaUcPeriod'
						| 'arabicParenBoth'
						| 'arabicParenR'
						| 'arabicPeriod'
						| 'arabicPlain'
						| 'romanLcParenBoth'
						| 'romanLcParenR'
						| 'romanLcPeriod'
						| 'romanUcParenBoth'
						| 'romanUcParenR'
						| 'romanUcPeriod'
					/**
					 * Number bullets start at
					 * @since v3.3.0
					 * @default 1
					 * @example 10 // numbered bullets start with 10
					 */
					numberStartAt?: number

					// DEPRECATED

					/**
					 * Bullet code (unicode)
					 * @deprecated v3.3.0 - use `characterCode`
					 */
					code?: string
					/**
					 * Margin between bullet and text
					 * @since v3.2.1
					 * @deplrecated v3.3.0 - use `indent`
					 */
					marginPt?: number
					/**
					 * Number to start with (only applies to type:number)
					 * @deprecated v3.3.0 - use `numberStartAt`
					 */
					startAt?: number
					/**
					 * Number type
					 * @deprecated v3.3.0 - use `numberType`
					 */
					style?: string
			  }
		/**
		 * Text color
		 * - `HexColor` or `ThemeColor`
		 * @example 'FF0000' // red
		 * @example 'pptxgen.SchemeColor.text1' // Text1 Theme Color
		 */
		color?: Color
		/**
		 * Font face name
		 * @example 'Arial' // Arial font
		 */
		fontFace?: string
		/**
		 * Font size
		 * @example 12 // Font size 12
		 */
		fontSize?: number
		/**
		 * Text highlight color (hex format)
		 * @example 'FFFF00' // yellow
		 */
		highlight?: HexColor
		/**
		 * italic style
		 * @default false
		 */
		italic?: boolean
		/**
		 * language
		 * - ISO 639-1 standard language code
		 * @default 'en-US' // english US
		 * @example 'fr-CA' // french Canadian
		 */
		lang?: string
		/**
		 * tab stops
		 * - PowerPoint: Paragraph > Tabs > Tab stop position
		 * @example [{ position:1 }, { position:3 }] // Set first tab stop to 1 inch, set second tab stop to 3 inches
		 */
		tabStops?: { position: number; alignment?: 'l' | 'r' | 'ctr' | 'dec' }[]
		/**
		 * underline properties
		 * - PowerPoint: Font > Color & Underline > Underline Style/Underline Color
		 * @default (none)
		 */
		underline?: {
			style?:
				| 'dash'
				| 'dashHeavy'
				| 'dashLong'
				| 'dashLongHeavy'
				| 'dbl'
				| 'dotDash'
				| 'dotDashHeave'
				| 'dotDotDash'
				| 'dotDotDashHeavy'
				| 'dotted'
				| 'dottedHeavy'
				| 'heavy'
				| 'none'
				| 'sng'
				| 'wavy'
				| 'wavyDbl'
				| 'wavyHeavy'
			color?: Color
		}
		/**
		 * vertical alignment
		 * @default 'top'
		 */
		valign?: VAlign
	}

	// image / media ==================================================================================
	export type MediaType = 'audio' | 'online' | 'video'

	export interface ImageProps extends PositionProps, DataOrPathProps {
		hyperlink?: HyperlinkProps
		placeholder?: string // 'body' | 'title' | etc.
		/**
		 * Image rotation (degrees)
		 * - range: -360 to 360
		 * @default 0
		 * @example 180 // rotate image 180 degrees
		 */
		rotate?: number
		/**
		 * Enable image rounding
		 * @default false
		 */
		rounding?: boolean
		/**
		 * Image sizing options
		 */
		sizing?: {
			/**
			 * Sizing type
			 */
			type: 'contain' | 'cover' | 'crop'
			/**
			 * Image width
			 */
			w: number
			/**
			 * Image height
			 */
			h: number
			/**
			 * Area horizontal position related to the image
			 * - Values: 0-n
			 * - `crop` only
			 */
			x?: number
			/**
			 * Area vertical position related to the image
			 * - Values: 0-n
			 * - `crop` only
			 */
			y?: number
		}
		/**
		 * Flip horizontally?
		 * @default false
		 */
		flipH?: boolean
		/**
		 * Flip vertical?
		 * @default false
		 */
		flipV?: boolean
	}
	/**
	 * Add media (audio/video) to slide
	 * @requires either `link` or `path`
	 */
	export interface MediaProps extends PositionProps, DataOrPathProps {
		/**
		 * Media type
		 * - Use 'online' to embed a YouTube video (only supported in recent versions of PowerPoint)
		 */
		type: MediaType
		/**
		 * video embed link
		 * - works with YouTube
		 * - other sites may not show correctly in PowerPoint
		 * @example 'https://www.youtube.com/embed/Dph6ynRVyUc' // embed a youtube video
		 */
		link?: string
		/**
		 * full or local path
		 * @example 'https://freesounds/simpsons/bart.mp3' // embed mp3 audio clip from server
		 * @example '/sounds/simpsons_haha.mp3' // embed mp3 audio clip from local directory
		 */
		path?: string
	}

	// shapes =========================================================================================

	export interface ShapeProps extends PositionProps {
		/**
		 * Horizontal alignment
		 * @default 'left'
		 */
		align?: HAlign
		/**
		 * Radius (only for pptx.shapes.PIE, pptx.shapes.ARC, pptx.shapes.BLOCK_ARC)
		 * - In the case of pptx.shapes.BLOCK_ARC you have to setup the arcThicknessRatio
		 * - values: [0-359, 0-359]
		 * @since v3.4.0
		 * @default [270, 0]
		 */
		angleRange?: [number, number]
		/**
		 * Radius (only for pptx.shapes.BLOCK_ARC)
		 * - You have to setup the angleRange values too
		 * - values: 0.0-1.0
		 * @since v3.4.0
		 * @default 0.5
		 */
		arcThicknessRatio?: number
		/**
		 * Shape fill color properties
		 * @example { color:'FF0000' } // hex string (red)
		 * @example { color:'pptx.SchemeColor.accent1' } // theme color Accent1
		 * @example { color:'0088CC', transparency:50 } // 50% transparent color
		 */
		fill?: ShapeFillProps
		/**
		 * Flip shape horizontally?
		 * @default false
		 */
		flipH?: boolean
		/**
		 * Flip shape vertical?
		 * @default false
		 */
		flipV?: boolean
		/**
		 * Add hyperlink to shape
		 * @example hyperlink: { url: "https://github.com/gitbrent/pptxgenjs", tooltip: "Visit Homepage" },
		 */
		hyperlink?: HyperlinkProps
		/**
		 * Line options
		 */
		line?: ShapeLineProps
		/**
		 * Radius (only for pptx.shapes.ROUNDED_RECTANGLE)
		 * - values: 0-180(TODO:values?)
		 * @default 0
		 */
		rectRadius?: number
		/**
		 * Image rotation (degrees)
		 * - range: -360 to 360
		 * @default 0
		 * @example 180 // rotate image 180 degrees
		 */
		rotate?: number
		/**
		 * Shadow options
		 * TODO: need new demo.js entry for shape shadow
		 */
		shadow?: ShadowProps
		/**
		 * Shape name
		 * - used instead of default "Shape N" name
		 * @since v3.3.0
		 * @example 'Antenna Design 9'
		 */
		shapeName?: string

		/**
		 * @depreacted v3.3.0
		 */
		lineSize?: number
		/**
		 * @depreacted v3.3.0
		 */
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		/**
		 * @depreacted v3.3.0
		 */
		lineHead?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @depreacted v3.3.0
		 */
		lineTail?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
	}

	// tables =========================================================================================

	export interface TableToSlidesProps extends TableProps {
		//_arrObjTabHeadRows?: TableRow[]
		//_masterSlide?: SlideLayout

		/**
		 * Add an image to slide(s) created during autopaging
		 */
		addImage?: { url: string; x: number; y: number; w?: number; h?: number }
		/**
		 * Add a shape to slide(s) created during autopaging
		 */
		addShape?: { shape: any; options: {} }
		/**
		 * Add a table to slide(s) created during autopaging
		 */
		addTable?: { rows: any[]; options: {} }
		/**
		 * Add a text object to slide(s) created during autopaging
		 */
		addText?: { text: any[]; options: {} }
		/**
		 * Whether to enable auto-paging
		 * - auto-paging creates new slides as content overflows a slide
		 * @default true
		 */
		autoPage?: boolean
		/**
		 * Auto-paging character weight
		 * - adjusts how many characters are used before lines wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
		 */
		autoPageCharWeight?: number
		/**
		 * Auto-paging line weight
		 * - adjusts how many lines are used before slides wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
		 */
		autoPageLineWeight?: number
		/**
		 * Whether to repeat head row(s) on new tables created by autopaging
		 * @since v3.3.0
		 * @default false
		 */
		autoPageRepeatHeader?: boolean
		/**
		 * The `y` location to use on subsequent slides created by autopaging
		 * @default (top margin of Slide)
		 */
		autoPageSlideStartY?: number
		/**
		 * Column widths (inches)
		 */
		colW?: number | number[]
		/**
		 * Master slide name
		 * - define a master slide to have your auto-paged slides have corporate design, etc.
		 * @see https://gitbrent.github.io/PptxGenJS/docs/masters.html
		 */
		masterSlideName?: string
		/**
		 * Slide margin
		 * - this margin will be across all slides created by auto-paging
		 */
		slideMargin?: Margin
		/**
		 * DEV TOOL: Verbose Mode (to console)
		 * - tell the library to provide an almost ridiculous amount of detail during auto-paging calculations
		 * @default false // obviously
		 */
		verbose?: boolean // Undocumented; shows verbose output

		/**
		 * @deprecated v3.3.0 - use `autoPageRepeatHeader`
		 */
		addHeaderToEach?: boolean
		/**
		 * @deprecated v3.3.0 - use `autoPageSlideStartY`
		 */
		newSlideStartY?: number
	}
	export interface TableCellProps extends TextBaseProps {
		/**
		 * Auto-paging character weight
		 * - adjusts how many characters are used before lines wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
		 */
		autoPageCharWeight?: number
		/**
		 * Auto-paging line weight
		 * - adjusts how many lines are used before slides wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
		 */
		autoPageLineWeight?: number
		/**
		 * Cell border
		 */
		border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
		/**
		 * Cell colspan
		 */
		colspan?: number
		/**
		 * Fill color
		 * @example { color:'FF0000' } // hex string (red)
		 * @example { color:'pptx.SchemeColor.accent1' } // theme color Accent1
		 * @example { color:'0088CC', transparency:50 } // 50% transparent color
		 * @example { type:'solid', color:'0088CC', alpha:50 } // ShapeFillProps object with 50% transparent
		 */
		fill?: ShapeFillProps
		/**
		 * Cell margin
		 * @default 0
		 */
		margin?: Margin
		/**
		 * Cell rowspan
		 */
		rowspan?: number
	}
	export interface TableProps extends PositionProps, TextBaseProps {
		//_arrObjTabHeadRows?: TableRow[]

		/**
		 * Whether to enable auto-paging
		 * - auto-paging creates new slides as content overflows a slide
		 * @default false
		 */
		autoPage?: boolean
		/**
		 * Auto-paging character weight
		 * - adjusts how many characters are used before lines wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // lines are longer (increases the number of characters that can fit on a given line)
		 */
		autoPageCharWeight?: number
		/**
		 * Auto-paging line weight
		 * - adjusts how many lines are used before slides wrap
		 * - range: -1.0 to 1.0
		 * @see https://gitbrent.github.io/PptxGenJS/docs/api-tables.html
		 * @default 0.0
		 * @example 0.5 // tables are taller (increases the number of lines that can fit on a given slide)
		 */
		autoPageLineWeight?: number
		/**
		 * Whether table header row(s) should be repeated on each new slide creating by autoPage.
		 * Use `autoPageHeaderRows` to designate how many rows comprise the table header (1+).
		 * @default false
		 * @since v3.3.0
		 */
		autoPageRepeatHeader?: boolean
		/**
		 * Number of rows that comprise table headers
		 * - required when `autoPageRepeatHeader` is set to true.
		 * @example 2 - repeats the first two table rows on each new slide created
		 * @default 1
		 * @since v3.3.0
		 */
		autoPageHeaderRows?: number
		/**
		 * The `y` location to use on subsequent slides created by autopaging
		 * @default (top margin of Slide)
		 */
		autoPageSlideStartY?: number
		/**
		 * Table border
		 * - single value is applied to all 4 sides
		 * - array of values in TRBL order for individual sides
		 */
		border?: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps]
		/**
		 * Width of table columns
		 * - single value is applied to every column equally based upon `w`
		 * - array of values in applied to each column in order
		 * @default columns of equal width based upon `w`
		 */
		colW?: number | number[]
		/**
		 * Cell background color
		 * @example { color:'FF0000' } // hex string (red)
		 * @example { color:'pptx.SchemeColor.accent1' } // theme color Accent1
		 * @example { color:'0088CC', transparency:50 } // 50% transparent color
		 */
		fill?: ShapeFillProps
		/**
		 * Cell margin
		 * - affects all table cells, is superceded by cell options
		 */
		margin?: Margin
		/**
		 * Height of table rows
		 * - single value is applied to every row equally based upon `h`
		 * - array of values in applied to each row in order
		 * @default rows of equal height based upon `h`
		 */
		rowH?: number | number[]

		/**
		 * @deprecated v3.3.0 - use `autoPageSlideStartY`
		 */
		newSlideStartY?: number
	}
	export interface TableCell {
		text?: string | TableCell[]
		options?: TableCellProps
	}
	export interface TableRowSlide {
		rows: TableRow[]
	}
	export type TableRow = TableCell[]

	// text ===========================================================================================
	export interface TextGlowProps {
		/**
		 * Border color (hex format)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * opacity (0.0 - 1.0)
		 * @example 0.5
		 * 50% opaque
		 */
		opacity: number
		/**
		 * size (points)
		 */
		size: number
	}

	export interface TextPropsOptions extends PositionProps, DataOrPathProps, TextBaseProps {
		baseline?: number
		/**
		 * Character spacing
		 */
		charSpacing?: number
		/**
		 * Text fit options
		 *
		 * MS-PPT > Format Shape > Shape Options > Text Box > "[unlabeled group]": [3 options below]
		 * - 'none' = Do not Autofit
		 * - 'shrink' = Shrink text on overflow
		 * - 'resize' = Resize shape to fit text
		 *
		 * **Note** 'shrink' and 'resize' only take effect after editing text/resize shape.
		 * Both PowerPoint and Word dynamically calculate a scaling factor and apply it when edit/resize occurs.
		 *
		 * There is no way for this library to trigger that behavior, sorry.
		 * @since v3.3.0
		 * @default "none"
		 */
		fit?: 'none' | 'shrink' | 'resize'
		/**
		 * Shape fill
		 * @example { color:'FF0000' } // hex string (red)
		 * @example { color:'pptx.SchemeColor.accent1' } // theme color Accent1
		 * @example { color:'0088CC', transparency:50 } // 50% transparent color
		 */
		fill?: ShapeFillProps
		/**
		 * Flip shape horizontally?
		 * @default false
		 */
		flipH?: boolean
		/**
		 * Flip shape vertical?
		 * @default false
		 */
		flipV?: boolean
		glow?: TextGlowProps
		hyperlink?: HyperlinkProps
		indentLevel?: number
		inset?: number
		isTextBox?: boolean
		line?: ShapeLineProps
		/**
		 * Line spacing (pt)
		 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Exactly"
		 * @example 28 // 28pt
		 */
		lineSpacing?: number
		/**
		 * line spacing multiple (percent)
		 * - range: 0.0-9.99
		 * - PowerPoint: Paragraph > Indents and Spacing > Line Spacing: > "Multiple"
		 * @example 1.5 // 1.5X line spacing
		 * @since v3.5.0
		 */
		lineSpacingMultiple?: number
		margin?: Margin
		outline?: { color: Color; size: number }
		paraSpaceAfter?: number
		paraSpaceBefore?: number
		placeholder?: string
		rotate?: number // (degree * 60,000)
		/**
		 * Whether to enable right-to-left mode
		 * @default false
		 */
		rtlMode?: boolean
		shadow?: ShadowProps
		shape?: SHAPE_NAME
		strike?: boolean | 'dblStrike' | 'sngStrike'
		subscript?: boolean
		superscript?: boolean
		valign?: VAlign
		vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
		/**
		 * Text wrap
		 * @since v3.3.0
		 * @default true
		 */
		wrap?: boolean

		/**
		 * Whather "Fit to Shape?" is enabled
		 * @deprecated v3.3.0 - use `fit`
		 */
		autoFit?: boolean
		/**
		 * Whather "Shrink Text on Overflow?" is enabled
		 * @deprecated v3.3.0 - use `fit`
		 */
		shrinkText?: boolean
		/**
		 * Dash type
		 * @deprecated v3.3.0 - use `line.dashType`
		 */
		lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * @deprecated v3.3.0 - use `line.beginArrowType`
		 */
		lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @deprecated v3.3.0 - use `line.width`
		 */
		lineSize?: number
		/**
		 * @deprecated v3.3.0 - use `line.endArrowType`
		 */
		lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	}
	export interface TextProps {
		text?: string
		options?: TextPropsOptions
	}

	// charts =========================================================================================
	export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
	export interface OptsChartData {
		index?: number
		labels?: string[]
		name?: string
		sizes?: number[]
		values?: number[]
		/**
		 * Override `chartColors`
		 */
		//color?: string // TODO: WIP: (Pull #727)
	}
	export interface OptsChartGridLine {
		/**
		 * Gridline color (hex)
		 * @example 'FF3399'
		 */
		color?: Color
		/**
		 * Gridline size (points)
		 */
		size?: number
		/**
		 * Gridline style
		 */
		style?: 'solid' | 'dash' | 'dot' | 'none'
	}
	export interface IChartMulti {
		type: CHART_NAME
		data: any[]
		options: {}
	}
	export interface IChartPropsBase {
		/**
		 * Axis position
		 */
		axisPos?: 'b' | 'l' | 'r' | 't'
		border?: BorderProps
		chartColors?: Color[]
		/**
		 * opacity (0.0 - 1.0)
		 * @example 0.5 // 50% opaque
		 */
		chartColorsOpacity?: number
		dataBorder?: BorderProps
		displayBlanksAs?: string
		fill?: Color
		invertedColors?: string
		lang?: string
		layout?: PositionProps
		shadow?: ShadowProps
		showLabel?: boolean
		showLeaderLines?: boolean
		showLegend?: boolean
		showPercent?: boolean
		showTitle?: boolean
		showValue?: boolean
		/**
		 * 3D perspecitve
		 * - range: 0-100
		 * @default 30
		 */
		v3DPerspective?: number
		v3DRAngAx?: boolean
		v3DRotX?: number
		v3DRotY?: number
	}
	export interface IChartPropsAxisCat {
		/**
		 * Multi-Chart prop: array of cat axes
		 */
		catAxes?: IChartPropsAxisCat[]
		catAxisBaseTimeUnit?: string
		catAxisHidden?: boolean
		catAxisLabelColor?: string
		catAxisLabelFontBold?: boolean
		catAxisLabelFontFace?: string
		catAxisLabelFontItalic?: boolean
		catAxisLabelFontSize?: number
		catAxisLabelFrequency?: string
		catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		catAxisLabelRotate?: number
		catAxisLineColor?: string
		catAxisLineShow?: boolean
		catAxisLineSize?: number
		catAxisLineStyle?: 'solid' | 'dash' | 'dot'
		catAxisMajorTickMark?: ChartAxisTickMark
		catAxisMajorTimeUnit?: string
		catAxisMajorUnit?: number
		catAxisMaxVal?: number
		catAxisMinorTickMark?: ChartAxisTickMark
		catAxisMinorTimeUnit?: string
		catAxisMinorUnit?: string
		catAxisMinVal?: number
		catAxisOrientation?: 'minMax'
		catAxisTitle?: string
		catAxisTitleColor?: Color
		catAxisTitleFontFace?: string
		catAxisTitleFontSize?: number
		catAxisTitleRotate?: number
		catGridLine?: OptsChartGridLine
		catLabelFormatCode?: string
		/**
		 * Whether data should use secondary category axis (instead of primary)
		 * @default false
		 */
		secondaryCatAxis?: boolean
		showCatAxisTitle?: boolean
	}
	export interface IChartPropsAxisSer {
		serAxisBaseTimeUnit?: string
		serAxisHidden?: boolean
		serAxisLabelColor?: string
		serAxisLabelFontBold?: boolean
		serAxisLabelFontFace?: string
		serAxisLabelFontItalic?: boolean
		serAxisLabelFontSize?: number
		serAxisLabelFrequency?: string
		serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		serAxisLineColor?: string
		serAxisLineShow?: boolean
		serAxisMajorTimeUnit?: string
		serAxisMajorUnit?: number
		serAxisMinorTimeUnit?: string
		serAxisMinorUnit?: number
		serAxisOrientation?: string
		serAxisTitle?: string
		serAxisTitleColor?: string
		serAxisTitleFontFace?: string
		serAxisTitleFontSize?: number
		serAxisTitleRotate?: number
		serGridLine?: OptsChartGridLine
		serLabelFormatCode?: string
		showSerAxisTitle?: boolean
	}
	export interface IChartPropsAxisVal {
		/**
		 * Whether data should use secondary value axis (instead of primary)
		 * @default false
		 */
		secondaryValAxis?: boolean
		showValAxisTitle?: boolean
		/**
		 * Multi-Chart prop: array of val axes
		 */
		valAxes?: IChartPropsAxisVal[]
		valAxisCrossesAt?: string | number
		valAxisDisplayUnit?: 'billions' | 'hundredMillions' | 'hundreds' | 'hundredThousands' | 'millions' | 'tenMillions' | 'tenThousands' | 'thousands' | 'trillions'
		valAxisDisplayUnitLabel?: boolean
		valAxisHidden?: boolean
		valAxisLabelColor?: string
		valAxisLabelFontBold?: boolean
		valAxisLabelFontFace?: string
		valAxisLabelFontItalic?: boolean
		valAxisLabelFontSize?: number
		valAxisLabelFormatCode?: string
		valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		valAxisLabelRotate?: number
		valAxisLineColor?: string
		valAxisLineShow?: boolean
		valAxisLineSize?: number
		valAxisLineStyle?: 'solid' | 'dash' | 'dot'
		/**
		 * PowerPoint: Format Axis > Axis Options > Logarithmic scale - Base
		 * - range: 2-99
		 * @since v3.5.0
		 */
		valAxisLogScaleBase?: number
		valAxisMajorTickMark?: ChartAxisTickMark
		valAxisMajorUnit?: number
		valAxisMaxVal?: number
		valAxisMinorTickMark?: ChartAxisTickMark
		valAxisMinVal?: number
		valAxisOrientation?: 'minMax'
		valAxisTitle?: string
		valAxisTitleColor?: Color
		valAxisTitleFontFace?: string
		valAxisTitleFontSize?: number
		valAxisTitleRotate?: number
		valGridLine?: OptsChartGridLine
		/**
		 * Value label format code
		 * - this also directs Data Table formatting
		 * @since v3.3.0
		 * @example '#%' // round percent
		 * @example '0.00%' // shows values as '0.00%'
		 * @example '$0.00' // shows values as '$0.00'
		 */
		valLabelFormatCode?: string
	}
	export interface IChartPropsChartBar {
		bar3DShape?: string
		barDir?: string
		barGapDepthPct?: number
		barGapWidthPct?: number
		barGrouping?: string
	}
	export interface IChartPropsChartDoughnut {
		dataNoEffects?: boolean
		holeSize?: number
	}
	export interface IChartPropsChartLine {
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		lineDataSymbol?: 'circle' | 'dash' | 'diamond' | 'dot' | 'none' | 'square' | 'triangle'
		lineDataSymbolLineColor?: string
		lineDataSymbolLineSize?: number
		lineDataSymbolSize?: number
		lineSize?: number
		lineSmooth?: boolean
	}
	export interface IChartPropsChartPie {
		dataNoEffects?: boolean
		/**
		 * MS-PPT > Format chart > Format Data Series > Series Options >  "Angle of first slice"
		 * - angle (degrees)
		 * - range: 0-359
		 * @since v3.4.0
		 * @default 0
		 */
		firstSliceAng?: number
	}
	export interface IChartPropsChartRadar {
		radarStyle?: 'standard' | 'marker' | 'filled'
	}
	export interface IChartPropsDataLabel {
		dataLabelBkgrdColors?: boolean
		dataLabelColor?: string
		dataLabelFontBold?: boolean
		dataLabelFontFace?: string
		dataLabelFontItalic?: boolean
		dataLabelFontSize?: number
		/**
		 * Data label format code
		 * @example '#%' // round percent
		 * @example '0.00%' // shows values as '0.00%'
		 * @example '$0.00' // shows values as '$0.00'
		 */
		dataLabelFormatCode?: string
		dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
		dataLabelPosition?: 'b' | 'bestFit' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd'
	}
	export interface IChartPropsDataTable {
		dataTableFontSize?: number
		/**
		 * Data table format code
		 * @since v3.3.0
		 * @example '#%' // round percent
		 * @example '0.00%' // shows values as '0.00%'
		 * @example '$0.00' // shows values as '$0.00'
		 */
		dataTableFormatCode?: string
		showDataTable?: boolean
		showDataTableHorzBorder?: boolean
		showDataTableKeys?: boolean
		showDataTableOutline?: boolean
		showDataTableVertBorder?: boolean
	}
	export interface IChartPropsLegend {
		legendColor?: string
		legendFontFace?: string
		legendFontSize?: number
		legendPos?: 'b' | 'l' | 'r' | 't' | 'tr'
	}
	export interface IChartPropsTitle extends TextBaseProps {
		title?: string
		titleAlign?: string
		titleBold?: boolean
		titleColor?: Color
		titleFontFace?: string
		titleFontSize?: number
		titlePos?: { x: number; y: number }
		titleRotate?: number
	}
	export interface IChartOpts
		extends IChartPropsAxisCat,
			IChartPropsAxisSer,
			IChartPropsAxisVal,
			IChartPropsBase,
			IChartPropsChartBar,
			IChartPropsChartDoughnut,
			IChartPropsChartLine,
			IChartPropsChartPie,
			IChartPropsChartRadar,
			IChartPropsDataLabel,
			IChartPropsDataTable,
			IChartPropsLegend,
			IChartPropsTitle,
			OptsChartGridLine,
			PositionProps {
		/**
		 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
		 * - PowerPoint: [right-click on a chart] > "Edit Alt Text..."
		 */
		altText?: string
	}
	export interface ISlideRelChart extends OptsChartData {
		type: CHART_NAME | IChartMulti[]
		opts: IChartOpts
		data: OptsChartData[]
		rId: number
		Target: string
		globalId: number
		fileName: string
	}

	// Core
	// ====
	export interface WriteBaseProps {
		/**
		 * Whether to compress export (can save substantial space, but takes a bit longer to export)
		 * @default false
		 * @since v3.5.0
		 */
		compression?: boolean
	}
	export interface WriteProps extends WriteBaseProps {
		/**
		 * Output type
		 * - values: 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array' | 'STREAM'
		 * @default 'blob'
		 */
		outputType?: WRITE_OUTPUT_TYPE
	}
	export interface WriteFileProps extends WriteBaseProps {
		/**
		 * Export file name
		 * @default 'Presentation.pptx'
		 */
		fileName?: string
	}
	export interface SectionProps {
		/**
		 * Section title
		 */
		title: string
		/**
		 * Section order - uses to add section at any index
		 * - values: 1-n
		 */
		order?: number
	}
	export interface PresLayout {
		/**
		 * Layout Name
		 * @example 'LAYOUT_WIDE'
		 */
		name: string
		width: number
		height: number
	}
	export interface SlideNumberProps extends PositionProps, TextBaseProps {
		margin?: Margin
	}
	export interface SlideMasterProps {
		/**
		 * Unique name for this master
		 */
		title: string
		margin?: Margin
		background?: BackgroundProps
		objects?: (
			| { chart: {} }
			| { image: {} }
			| { line: {} }
			| { rect: {} }
			| { text: TextProps }
			| {
					placeholder: {
						options: PlaceholderProps
						/**
						 * Text to be shown in placeholder (shown until user focuses textbox or adds text)
						 * - Leave blank to have powerpoint show default phrase (ex: "Click to add title")
						 */
						text?: string
					}
			  }
		)[]
		slideNumber?: SlideNumberProps

		/**
		 * @deprecated v3.3.0 - use `background`
		 */
		bkgd?: string | BackgroundProps
	}
	export interface ObjectOptions extends ImageProps, PositionProps, ShapeProps, TableCellProps, TextPropsOptions {
		cx?: Coord
		cy?: Coord
		margin?: Margin
		colW?: number | number[] // table
		rowH?: number | number[] // table
	}
	export interface PresSlide {
		addChart: Function
		addImage: Function
		addMedia: Function
		addNotes: Function
		addShape: Function
		addTable: Function
		addText: Function

		/**
		 * Background color or image (`Color` | `path` | `data`)
		 * @example {color: 'FF3399'} - hex fill color
		 * @example {color: 'FF3399', transparency:50} - hex fill color with transparency of 50%
		 * @example {path: 'https://onedrives.com/myimg.png`} - retrieve image via URL
		 * @example {path: '/home/gitbrent/images/myimg.png`} - retrieve image via local path
		 * @example {data: 'image/png;base64,iVtDaDrF[...]='} - base64 string
		 * @since v3.3.0
		 */
		background?: BackgroundProps
		/**
		 * Default text color (hex format)
		 * @example 'FF3399'
		 * @default '000000' (DEF_FONT_COLOR)
		 */
		color?: HexColor
		/**
		 * Whether slide is hidden
		 * @default false
		 */
		hidden?: boolean
		/**
		 * Slide number options
		 */
		slideNumber?: SlideNumberProps
	}
	export interface AddSlideProps {
		masterName?: string // TODO: 20200528: rename to "masterTitle" (createMaster uses `title` so lets be consistent)
		sectionTitle?: string
	}
	export interface PresentationProps {
		author: string
		company: string
		layout: string
		masterSlide: PresSlide
		/**
		 * Presentation's layout
		 * read-only
		 */
		presLayout: PresLayout
		revision: string
		/**
		 * Whether to enable right-to-left mode
		 * @default false
		 */
		rtlMode: boolean
		subject: string
		title: string
	}

	// LAST: Slide
	/**
	 * `slide.d.ts`
	 */
	export class Slide {
		/**
		 * Background color or image (`Color` | `path` | `data`)
		 * @example {color: 'FF3399'} - hex fill color
		 * @example {color: 'FF3399', transparency: 50} - hex fill color with transparency of 50%
		 * @example {path: 'https://onedrives.com/myimg.png`} - retrieve image via URL
		 * @example {path: '/home/gitbrent/images/myimg.png`} - retrieve image via local path
		 * @example {data: 'image/png;base64,iVtDaDrF[...]='} - base64 string
		 * @since 3.3.0
		 */
		background: BackgroundProps
		/**
		 * Default text color (hex format)
		 * @example 'FF3399'
		 * @default '000000' (DEF_FONT_COLOR)
		 */
		color: HexColor
		/**
		 * Whether slide is hidden
		 * @default false
		 */
		hidden: boolean
		/**
		 * Slide number options
		 */
		slideNumber: SlideNumberProps
		/**
		 * Add chart to Slide
		 * @param {CHART_NAME|IChartMulti[]} type - chart type
		 * @param {object[]} data - data object
		 * @param {IChartOpts} options - chart options
		 * @return {Slide} this Slide
		 * @type {Function}
		 */
		addChart(type: CHART_NAME | IChartMulti[], data: any[], options?: IChartOpts): Slide
		/**
		 * Add image to Slide
		 * @param {ImageProps} options - image options
		 * @return {Slide} this Slide
		 */
		addImage(options: ImageProps): Slide
		/**
		 * Add media (audio/video) to Slide
		 * @param {MediaProps} options - media options
		 * @return {Slide} this Slide
		 */
		addMedia(options: MediaProps): Slide
		/**
		 * Add speaker notes to Slide
		 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
		 * @param {string} notes - notes to add to slide
		 * @return {Slide} this Slide
		 */
		addNotes(notes: string): Slide
		/**
		 * Add shape to Slide
		 * @param {SHAPE_NAME} shapeName - shape name
		 * @param {ShapeProps} options - shape options
		 * @return {Slide} this Slide
		 */
		addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide
		/**
		 * Add table to Slide
		 * @param {TableRow[]} tableRows - table rows
		 * @param {TableProps} options - table options
		 * @return {Slide} this Slide
		 */
		addTable(tableRows: TableRow[], options?: TableProps): Slide
		/**
		 * Add text to Slide
		 * @param {string|TextProps[]} text - text string or complex object
		 * @param {TextPropsOptions} options - text options
		 * @return {Slide} this Slide
		 */
		addText(text: string | TextProps[], options?: TextPropsOptions): Slide

		/**
		 * Background color
		 * @deprecated in 3.3.0 - use `background` instead
		 */
		bkgd: string
	}
}
