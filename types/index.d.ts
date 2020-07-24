// Type definitions for pptxgenjs 3.3.0
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
	readonly presLayout: PptxGenJS.ILayout
	readonly AlignH: typeof PptxGenJS.AlignH
	readonly AlignV: typeof PptxGenJS.AlignV
	readonly ChartType: typeof PptxGenJS.ChartType
	readonly OutputType: typeof PptxGenJS.OutputType
	readonly SchemeColor: typeof PptxGenJS.SchemeColor
	readonly ShapeType: typeof PptxGenJS.ShapeType

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
	 */
	rtlMode: boolean

	// Presentation Metadata
	author: string
	company: string
	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PowerPoint will throw errors upon opening!)
	 */
	revision: string
	subject: string
	title: string

	// Methods

	/**
	 * Export the current Presentation to stream
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file stream
	 */
	stream(): Promise<string | ArrayBuffer | Blob | Uint8Array>
	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {JSZIP_OUTPUT_TYPE} outputType - 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file content in selected type
	 */
	write(outputType: PptxGenJS.JSZIP_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Uint8Array>
	/**
	 * Export the current Presentation. Writes file to local file system if `fs` exists, otherwise, initiates download in browsers
	 * @param {string} exportName - file name
	 * @returns {Promise<string>} the presentation name
	 */
	writeFile(exportName?: string): Promise<string>
	/**
	 * Add a new Section to Presentation
	 * @param {ISectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(section: PptxGenJS.ISectionProps): void
	/**
	 * Add a new Slide to Presentation
	 * @param {IAddSlideOptions} options - slide options
	 * @returns {Slide} the new Slide
	 */
	addSlide(options?: PptxGenJS.IAddSlideOptions): PptxGenJS.Slide
	/**
	 * Add a new Slide to Presentation
	 * @param {string} masterName - master slide name
	 * @returns {Slide} the new Slide
	 * @deprecated use `addSlide(IAddSlideOptions)`
	 */
	addSlide(masterName?: string): PptxGenJS.Slide
	/**
	 * Create a custom Slide Layout in any size
	 * @param {ILayoutProps} layout - an object with user-defined w/h
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: PptxGenJS.ILayoutProps): void
	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {ISlideMasterOptions} slideMasterOpts - layout definition
	 */
	defineSlideMaster(slideMasterOpts: PptxGenJS.ISlideMasterOptions): void
	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} tabEleId - HTMLElementID of the table
	 * @param {TableToSlidesOpts} inOpts - array of options (e.g.: tabsize)
	 */
	tableToSlides(tableElementId: string, opts?: PptxGenJS.TableToSlidesOpts): void
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

	// `core-interfaces.d.ts`
	// import { CHART_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, PLACEHOLDER_TYPES, SHAPE_NAME } from './core-enums'
	export type JSZIP_OUTPUT_TYPE = 'arraybuffer' | 'base64' | 'binarystring' | 'blob' | 'nodebuffer' | 'uint8array'
	export type CHART_NAME = 'area' | 'bar' | 'bar3D' | 'bubble' | 'doughnut' | 'line' | 'pie' | 'radar' | 'scatter'
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
	export enum PLACEHOLDER_TYPES {
		'title' = 'title',
		'body' = 'body',
		'image' = 'pic',
		'chart' = 'chart',
		'table' = 'tbl',
		'media' = 'media',
	}
	export type SHAPE_NAME =
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
		| 'wedgeRoundRectCallout'
		| 'bentArrow'
		| 'bentUpArrow'
		| 'bevel'
		| 'blockArc'
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
		| 'plus'
		| 'cube'
		| 'curvedDownArrow'
		| 'ellipseRibbon'
		| 'curvedLeftArrow'
		| 'curvedRightArrow'
		| 'curvedUpArrow'
		| 'ellipseRibbon2'
		| 'decagon'
		| 'diagStripe'
		| 'diamond'
		| 'dodecagon'
		| 'donut'
		| 'bracePair'
		| 'bracketPair'
		| 'doubleWave'
		| 'downArrow'
		| 'downArrowCallout'
		| 'ribbon'
		| 'irregularSeal1'
		| 'irregularSeal2'
		| 'flowChartAlternateProcess'
		| 'flowChartPunchedCard'
		| 'flowChartCollate'
		| 'flowChartConnector'
		| 'flowChartInputOutput'
		| 'flowChartDecision'
		| 'flowChartDelay'
		| 'flowChartMagneticDrum'
		| 'flowChartDisplay'
		| 'flowChartDocument'
		| 'flowChartExtract'
		| 'flowChartInternalStorage'
		| 'flowChartMagneticDisk'
		| 'flowChartManualInput'
		| 'flowChartManualOperation'
		| 'flowChartMerge'
		| 'flowChartMultidocument'
		| 'flowChartOfflineStorage'
		| 'flowChartOffpageConnector'
		| 'flowChartOr'
		| 'flowChartPredefinedProcess'
		| 'flowChartPreparation'
		| 'flowChartProcess'
		| 'flowChartPunchedTape'
		| 'flowChartMagneticTape'
		| 'flowChartSort'
		| 'flowChartOnlineStorage'
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
		| 'horizontalScroll'
		| 'triangle'
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
		| 'borderCallout1'
		| 'accentCallout1'
		| 'accentBorderCallout1'
		| 'callout1'
		| 'borderCallout2'
		| 'accentCallout2'
		| 'accentBorderCallout2'
		| 'callout2'
		| 'borderCallout3'
		| 'accentCallout3'
		| 'accentBorderCallout3'
		| 'callout3'
		| 'borderCallout3'
		| 'accentCallout3'
		| 'accentBorderCallout3'
		| 'callout3'
		| 'line'
		| 'lineInv'
		| 'mathDivide'
		| 'mathEqual'
		| 'mathMinus'
		| 'mathMultiply'
		| 'mathNotEqual'
		| 'mathPlus'
		| 'moon'
		| 'nonIsoscelesTrapezoid'
		| 'notchedRightArrow'
		| 'noSmoking'
		| 'octagon'
		| 'ellipse'
		| 'wedgeEllipseCallout'
		| 'parallelogram'
		| 'homePlate'
		| 'pie'
		| 'pieWedge'
		| 'plaque'
		| 'plaqueTabs'
		| 'quadArrow'
		| 'quadArrowCallout'
		| 'rect'
		| 'wedgeRectCallout'
		| 'pentagon'
		| 'rightArrow'
		| 'rightArrowCallout'
		| 'rightBrace'
		| 'rightBracket'
		| 'rtTriangle'
		| 'roundRect'
		| 'wedgeRoundRectCallout'
		| 'round1Rect'
		| 'round2DiagRect'
		| 'round2SameRect'
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
		| 'upArrow'
		| 'upArrowCallout'
		| 'upDownArrow'
		| 'upDownArrowCallout'
		| 'ribbon2'
		| 'uturnArrow'
		| 'verticalScroll'
		| 'wave'

	export interface ISectionProps {
		title: string
		/**
		 * Section order [index] (1-n)
		 */
		order?: number
	}
	export interface ILayoutProps {
		name: string
		width: number
		height: number
	}
	export interface ILayout {
		name: string
		width?: number
		height?: number
	}
	export interface ISlideLayout {
		presLayout: ILayout
		name: string
		number: number
		bkgd?: string
		bkgdImgRid?: number
		slide?: {
			back: string
			bkgdImgRid?: number
			color: string
			hidden?: boolean
		}
		data: ISlideObject[]
		rels: ISlideRel[]
		relsChart: ISlideRelChart[]
		relsMedia: ISlideRelMedia[]
		margin?: Margin
		slideNumberObj?: ISlideNumber
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
	export interface ISlideRel {
		type: SLIDE_OBJECT_TYPES
		Target: string
		fileName?: string
		data: any[] | string
		opts?: IChartOpts
		path?: string
		extn?: string
		globalId?: number
		rId: number
	}
	export interface ISlideRelMedia {
		type: string
		opts?: MediaOpts
		path?: string
		extn?: string
		data?: string | ArrayBuffer
		isSvgPng?: boolean
		svgSize?: {
			w: number
			h: number
		}
		rId: number
		Target: string
	}
	export interface ISlideMasterOptions {
		title: string
		height?: number
		width?: number
		margin?: Margin
		background?: BkgdOpts
		bkgd?: string | BkgdOpts // @deprecated
		objects?: (
			| {
					chart: {} // TODO: IChartOptions (?)
			  }
			| {
					image: {} // TODO: IImageOptions (?)
			  }
			| {
					line: {} // TODO: ShapeOptions (?)
			  }
			| {
					rect: {} // TODO: ShapeOptions (?)
			  }
			| {
					text: {
						text: string
						options?: ITextOpts
					}
			  }
			| {
					placeholder: {
						text: string
						options?: ISlideMstrObjPlchldrOpts
					}
			  }
		)[]
		slideNumber?: ISlideNumber
	}
	export interface ISlideMstrObjPlchldrOpts {
		name: string
		type: PLACEHOLDER_TYPES
		x: Coord
		y: Coord
		w: Coord
		h: Coord
	}
	export interface IAddSlideOptions {
		masterName?: string
		sectionTitle?: string
	}
	export interface ISlide {
		addChart: Function
		addImage: Function
		addMedia: Function
		addNotes: Function
		addShape: Function
		addTable: Function
		addText: Function
		bkgd?: string
		color?: string
		hidden?: boolean
		slideNumber?: ISlideNumber
	}

	export interface OptsChartData {
		index?: number
		name?: string
		labels?: string[]
		values?: number[]
		sizes?: number[]
	}

	export interface IObjectOptions extends ShapeOptions, TableCellOpts, ITextOpts {
		x?: Coord
		y?: Coord
		cx?: Coord
		cy?: Coord
		w?: Coord
		h?: Coord
		margin?: Margin
		colW?: number | number[]
		rowH?: number | number[]
		sizing?: {
			type?: string
			x?: number
			y?: number
			w?: number
			h?: number
		}
		rounding?: string
		placeholderIdx?: number
		placeholderType?: PLACEHOLDER_TYPES
	}
	export interface ISlideObject {
		type: SLIDE_OBJECT_TYPES
		options?: IObjectOptions
		text?: string | IText[]
		arrTabRows?: TableCell[][]
		chartRid?: number
		image?: string
		imageRid?: number
		hyperlink?: HyperLink
		media?: string
		mtype?: MediaType
		mediaRid?: number
		shape?: SHAPE_NAME
	}

	/**
	 * Coordinate number - either:
	 * - Inches
	 * - Percentage
	 *
	 * @example 10.25
	 * coordinate in inches
	 * @example '75%'
	 * coordinate in percentage of slide size
	 */
	export type Coord = number | string
	/**
	 * Color in Hex format
	 * @example 'FF3399'
	 */
	export type HexColor = string
	export interface OptsDataOrPath {
		/**
		 * URL or relative path
		 *
		 * @example 'https://onedrives.com/myimg.png`
		 * retrieve image via URL
		 * @example '/home/gitbrent/images/myimg.png`
		 * retrieve image via local path
		 */
		path?: string
		/**
		 * base64-encoded string
		 * - Useful for avoiding potential path/server issues
		 *
		 * @example 'image/png;base64,iVtDafDrBF[...]='
		 * adds a pre-encoded image
		 */
		data?: string
	}
	export interface PositionOptions {
		/**
		 * Horizontal position
		 * - inches or percentage
		 * @example 10.25
		 * position in inches
		 * @example '75%'
		 * position as percentage of slide size
		 */
		x?: Coord
		/**
		 * Vertical position
		 * - inches or percentage
		 * @example 10.25
		 * position in inches
		 * @example '75%'
		 * position as percentage of slide size
		 */
		y?: Coord
		/**
		 * Height
		 * - inches or percentage
		 * @example 10.25
		 * height in inches
		 * @example '75%'
		 * height as percentage of slide size
		 */
		h?: Coord
		/**
		 * Width
		 * - inches or percentage
		 * @example 10.25
		 * width in inches
		 * @example '75%'
		 * width as percentage of slide size
		 */
		w?: Coord
	}
	export interface IBorderOptions {
		/**
		 * Border color (hex format)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * Border size (points)
		 */
		pt?: number
		/**
		 * Border type
		 */
		type?: 'none' | 'dash' | 'solid'
	}
	export interface IShadowOptions {
		/**
		 * shadow type
		 */
		type: 'outer' | 'inner' | 'none'
		/**
		 * opacity (0.0 - 1.0)
		 * @example 0.5
		 * 50% opaque
		 */
		opacity: number
		/**
		 * blue (points)
		 * - range: 0-100
		 */
		blur?: number
		/**
		 * angle (degrees)
		 * - range: 0-359
		 */
		angle: number
		/**
		 * shadow offset (points)
		 * - range: 0-200
		 */
		offset?: number
		/**
		 * shadow color (hex format)
		 * @example 'FF3399'
		 */
		color?: HexColor
	}
	export interface IGlowOptions {
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
	export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
	export type Color = HexColor | ThemeColor
	export type Margin = number | [number, number, number, number]
	export type HAlign = 'left' | 'center' | 'right' | 'justify'
	export type VAlign = 'top' | 'middle' | 'bottom'
	export type MediaType = 'audio' | 'online' | 'video'
	export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
	export type ShapeFill = {
		/**
		 * Fill type
		 * @deprecated 'solid'
		 */
		type?: 'none' | 'solid'
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
		 * Transparency (percent)
		 * @deprecated v3.3.0 - use `transparency`
		 */
		alpha?: number
	}
	export interface ShapeLine extends ShapeFill {
		/**
		 * Line size (pt)
		 * @default 1
		 */
		size?: number
		/**
		 * Dash type
		 * @default 'solid'
		 */
		dashType?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * Begin arrow type
		 */
		beginArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * End arrow type
		 */
		endArrowType?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'

		/**
		 * Dash type
		 * @deprecated v3.3.0 - use `dashType`
		 */
		lineDash?: 'solid' | 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'sysDash' | 'sysDot'
		/**
		 * @deprecated v3.3.0 - use `arrowTypeBegin`
		 */
		lineHead?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
		/**
		 * @deprecated v3.3.0 - use `arrowTypeEnd`
		 */
		lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
	}
	export type HyperLink = {
		slide?: number
		tooltip?: string
		url?: string
	}
	export interface BkgdOpts extends OptsDataOrPath {
		/**
		 * Color in Hex format
		 * @example 'FF3399'
		 */
		fill?: HexColor
	}
	export type TextOptions = {
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
					 * Bullet code (unicode)
					 * @deprecated 3.3.0
					 */
					code?: string
					/**
					 * Bullet character code (unicode)
					 * @since 3.3.0
					 * @example { code: '25BA' } // 'BLACK RIGHT-POINTING POINTER' (U+25BA)
					 */
					characterCode?: string
					/**
					 * Indentation (space between bullet and text) (points)
					 * @since 3.3.0
					 * @example { margin: 10 } // 10 points between bullet and text
					 */
					indent?: number // TODO: new!
					/**
					 * Margin between bullet and text
					 * @since 3.2.1
					 * @deplrecated 3.3.0
					 */
					marginPt?: number
					/**
					 * Number type
					 * @since 3.3.0
					 * @example romanLcParenR // roman numerals lower-case with paranthesis right
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
					 * @since 3.3.0
					 * @example { numberStartAt: 10 } // numbered bullets start with 10.
					 */
					numberStartAt?: number
					/**
					 * Number to start with (only applies to type:number)
					 * @deprecated 3.3.0 - use `numberStartAt` instead
					 */
					startAt?: number
					/**
					 * Number type
					 * @deprecated 3.3.0 use `numberType` instead
					 */
					style?: string
			  }
		/**
		 * Text color
		 * - `HexColor` or `ThemeColor`
		 * @example 'FF0000' // red
		 * @example 'pptx.SchemeColor.text1' // Text1 Theme Color
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
		 * vertical alignment
		 * @default 'top'
		 */
		valign?: VAlign
	}

	// slideNumber
	export interface ISlideNumber extends PositionOptions, TextOptions {
		align?: HAlign
		color?: string
	}

	// addChart
	export type OptsChartGridLine = {
		/**
		 * Gridline color (hex)
		 * @example 'FF3399'
		 */
		color?: HexColor
		/**
		 * Gridline size (points)
		 */
		size?: number
		/**
		 * Gridline style
		 */
		style?: 'solid' | 'dash' | 'dot' | 'none'
	}
	export interface IChartTitleOpts extends TextOptions {
		color?: Color
		rotate?: number
		title: string
		titleAlign?: string
		titlePos?: { x: number; y: number }
	}
	export interface IChartMulti {
		type: CHART_NAME
		data: any[]
		options: {}
	}
	export interface IChartPropsBase {
		axisPos?: string
		border?: IBorderOptions
		chartColors?: string[]
		chartColorsOpacity?: number
		dataBorder?: IBorderOptions
		displayBlanksAs?: string
		fill?: string
		invertedColors?: string
		lang?: string
		layout?: PositionOptions
		shadow?: IShadowOptions
		showLabel?: boolean
		showLeaderLines?: boolean
		showLegend?: boolean
		showPercent?: boolean
		showTitle?: boolean
		showValue?: boolean
		v3DPerspective?: number
		v3DRAngAx?: boolean
		v3DRotX?: number
		v3DRotY?: number
	}
	export interface IChartPropsAxisCat {
		catAxes?: number[]
		catAxisBaseTimeUnit?: string
		catAxisHidden?: boolean
		catAxisLabelColor?: string
		catAxisLabelFontBold?: boolean
		catAxisLabelFontFace?: string
		catAxisLabelFontSize?: number
		catAxisLabelFrequency?: string
		catAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		catAxisLabelRotate?: number
		catAxisLineShow?: boolean
		catAxisMajorTickMark?: ChartAxisTickMark
		catAxisMajorTimeUnit?: string
		catAxisMajorUnit?: number
		catAxisMaxVal?: number
		catAxisMinorTickMark?: ChartAxisTickMark
		catAxisMinorTimeUnit?: string
		catAxisMinorUnit?: string
		catAxisMinVal?: number
		catAxisOrientation?: 'minMax' | 'minMax'
		catAxisTitle?: string
		catAxisTitleColor?: string
		catAxisTitleFontFace?: string
		catAxisTitleFontSize?: number
		catAxisTitleRotate?: number
		catGridLine?: OptsChartGridLine
		catLabelFormatCode?: string
		showCatAxisTitle?: boolean
	}
	export interface IChartPropsAxisSer {
		serAxisBaseTimeUnit?: string
		serAxisHidden?: boolean
		serAxisLabelColor?: string
		serAxisLabelFontFace?: string
		serAxisLabelFontSize?: string
		serAxisLabelFrequency?: string
		serAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
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
		showValAxisTitle?: boolean
		valAxes?: number[]
		valAxisCrossesAt?: string | number
		valAxisDisplayUnit?: 'billions' | 'hundredMillions' | 'hundreds' | 'hundredThousands' | 'millions' | 'tenMillions' | 'tenThousands' | 'thousands' | 'trillions'
		valAxisDisplayUnitLabel?: boolean
		valAxisHidden?: boolean
		valAxisLabelColor?: string
		valAxisLabelFontBold?: boolean
		valAxisLabelFontFace?: string
		valAxisLabelFontSize?: number
		valAxisLabelFormatCode?: string
		valAxisLabelPos?: 'none' | 'low' | 'high' | 'nextTo'
		valAxisLabelRotate?: number
		valAxisLineShow?: boolean
		valAxisMajorTickMark?: ChartAxisTickMark
		valAxisMajorUnit?: number
		valAxisMaxVal?: number
		valAxisMinorTickMark?: ChartAxisTickMark
		valAxisMinVal?: number
		valAxisOrientation?: 'minMax' | 'minMax'
		valAxisTitle?: string
		valAxisTitleColor?: string
		valAxisTitleFontFace?: string
		valAxisTitleFontSize?: number
		valAxisTitleRotate?: number
		valGridLine?: OptsChartGridLine
	}
	export interface IChartPropsChartBar {
		bar3DShape?: string
		barDir?: string
		barGapDepthPct?: number
		barGapWidthPct?: number
		barGrouping?: string
		valueBarColors?: string[]
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
	}
	export interface IChartPropsChartRadar {
		radarStyle?: 'standard' | 'marker' | 'filled'
	}
	export interface IChartPropsDataLabel {
		dataLabelBkgrdColors?: boolean
		dataLabelColor?: string
		dataLabelFontBold?: boolean
		dataLabelFontFace?: string
		dataLabelFontSize?: number
		dataLabelFormatCode?: string
		dataLabelFormatScatter?: 'custom' | 'customXY' | 'XY'
		dataLabelPosition?: 'b' | 'bestFit' | 'ctr' | 'l' | 'r' | 't' | 'inEnd' | 'outEnd' | 'bestFit'
	}
	export interface IChartPropsDataTable {
		dataTableFontSize?: number
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
	export interface IChartPropsTitle {
		title?: string
		titleAlign?: string
		titleColor?: string
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
			PositionOptions {}

	// addImage
	export interface ImageOpts extends PositionOptions, OptsDataOrPath {
		hyperlink?: HyperLink
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
			x?: number
			y?: number
		}
	}

	// addMedia
	/**
	 * Add media (audio/video) to slide
	 * @requires either `link` or `path`
	 */
	export interface MediaOpts extends PositionOptions, OptsDataOrPath {
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

	// addShape
	export interface ShapeOptions extends PositionOptions {
		align?: HAlign
		fill?: ShapeFill
		flipH?: boolean
		flipV?: boolean
		lineSize?: number
		lineDash?: 'dash' | 'dashDot' | 'lgDash' | 'lgDashDot' | 'lgDashDotDot' | 'solid' | 'sysDash' | 'sysDot'
		lineHead?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		lineTail?: 'arrow' | 'diamond' | 'none' | 'oval' | 'stealth' | 'triangle'
		line?: ShapeLine
		rectRadius?: number
		rotate?: number
		shadow?: IShadowOptions
	}

	// addTable & tableToSlides
	export interface TableToSlidesOpts extends TableOptions {
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
		 * @since 3.3.0
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
		 * @deprecated 3.3.0 - use `autoPageRepeatHeader`
		 */
		addHeaderToEach?: boolean
		/**
		 * @deprecated 3.3.0 - use `autoPageSlideStartY`
		 */
		newSlideStartY?: number
	}
	export interface TableCellOpts extends TextOptions {
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
		border?: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
		/**
		 * Cell colspan
		 */
		colspan?: number
		/**
		 * Fill color
		 * @example 'FF0000' // hex string (red)
		 * @example 'pptx.SchemeColor.accent1' // theme color Accent1
		 * @example { type:'solid', color:'0088CC', alpha:50 } // ShapeFill object with 50% transparent
		 */
		fill?: ShapeFill
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
	export interface TableOptions extends PositionOptions, TextOptions {
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
		border?: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
		/**
		 * Width of table columns
		 * - single value is applied to every column equally based upon `w`
		 * - array of values in applied to each column in order
		 * @default columns of equal width based upon `w`
		 */
		colW?: number | number[]
		/**
		 * Cell background color
		 */
		fill?: ShapeFill
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
		 * @deprecated 3.3.0 - use `autoPageSlideStartY`
		 */
		newSlideStartY?: number
	}
	export interface TableCell {
		text?: string | TableCell[]
		options?: TableCellOpts
	}
	export interface TableRowSlide {
		rows: TableRow[]
	}
	export type TableRow = number[] | string[] | TableCell[]

	// addText
	export interface ITextOpts extends PositionOptions, OptsDataOrPath, TextOptions {
		autoFit?: boolean
		bodyProp?: {
			// Note: Many of these duplicated as user options are transformed to bodyProp options for XML processing
			autoFit?: boolean
			align?: TEXT_HALIGN
			anchor?: TEXT_VALIGN
			lIns?: number
			rIns?: number
			tIns?: number
			bIns?: number
			vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
			wrap?: boolean
		}
		charSpacing?: number
		fill?: ShapeFill
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
		glow?: IGlowOptions
		hyperlink?: HyperLink
		indentLevel?: number
		inset?: number
		isTextBox?: boolean
		line?: ShapeLine
		lineIdx?: number
		lineSize?: number
		lineSpacing?: number
		margin?: Margin
		outline?: { color: Color; size: number }
		paraSpaceAfter?: number
		paraSpaceBefore?: number
		placeholder?: string
		rotate?: number // (degree * 60,000)
		rtlMode?: boolean
		shadow?: IShadowOptions
		shape?: SHAPE_NAME
		shrinkText?: boolean
		strike?: boolean
		subscript?: boolean
		superscript?: boolean
		underline?: boolean
		valign?: VAlign
		vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
		wrap?: boolean
	}
	export interface IText {
		text: string
		options?: ITextOpts
	}

	/**
	 * `slide.d.ts`
	 */
	export class Slide {
		/**
		 * Background color
		 * @type {string}
		 * @deprecated in v3.3.0 - use `background` instead
		 */
		bkgd: string
		/**
		 * Background color or image
		 * @type {BkgdOpts}
		 * @example `background: {fill:'FF0000'}
		 * @example `background: {data:'image/png;base64,ABC[...]123'}`
		 * @example `background: {path:'https://some.url/image.jpg'}`
		 */
		background: BkgdOpts
		/**
		 * Default font color
		 */
		color: HexColor
		hidden: boolean
		slideNumber: ISlideNumber
		/**
		 * Add chart to Slide
		 * @param {CHART_NAME|IChartMulti[]} type - chart type
		 * @param {object[]} data - data object
		 * @param {IChartOpts} options - chart options
		 * @return {Slide} this Slide
		 */
		addChart(type: CHART_NAME | IChartMulti[], data: any[], options?: IChartOpts): Slide
		/**
		 * Add image to Slide
		 * @param {ImageOpts} options - image options
		 * @return {Slide} this Slide
		 */
		addImage(options: ImageOpts): Slide
		/**
		 * Add media (audio/video) to Slide
		 * @param {MediaOpts} options - media options
		 * @return {Slide} this Slide
		 */
		addMedia(options: MediaOpts): Slide
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
		 * @param {ShapeOptions} options - shape options
		 * @return {Slide} this Slide
		 */
		addShape(shapeName: SHAPE_NAME, options?: ShapeOptions): Slide
		/**
		 * Add table to Slide
		 * @param {TableRow[]} tableRows - table rows
		 * @param {TableOptions} options - table options
		 * @return {Slide} this Slide
		 */
		addTable(tableRows: TableRow[], options?: TableOptions): Slide
		/**
		 * Add text to Slide
		 * @param {string|IText[]} text - text string or complex object
		 * @param {ITextOpts} options - text options
		 * @return {Slide} this Slide
		 */
		addText(text: string | IText[], options?: ITextOpts): Slide
	}
}
