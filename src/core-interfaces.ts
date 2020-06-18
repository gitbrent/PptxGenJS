/**
 * PptxGenJS Interfaces
 */

import { CHART_NAME, PLACEHOLDER_TYPES, SHAPE_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN } from './core-enums'

// Core Types
// ==========

/**
 * Coordinate number - either:
 * - Inches
 * - Percentage
 *
 * @example 10.25
 * coordinate in inches
 * @example '75%'
 * coordinate as percentage of slide size
 */
export type Coord = number | string
export type PositionOptions = {
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
/**
 * Either `data` or `path` is required
 */
export type OptsDataOrPath = {
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
export interface BkgdOpts extends OptsDataOrPath {
	/**
	 * Color (hex format)
	 * @example 'FF3399'
	 */
	fill?: HexColor
}
/**
 * Color in Hex format
 * @example 'FF3399'
 */
export type HexColor = string
export type ThemeColor = 'tx1' | 'tx2' | 'bg1' | 'bg2' | 'accent1' | 'accent2' | 'accent3' | 'accent4' | 'accent5' | 'accent6'
export type Color = HexColor | ThemeColor
export type Margin = number | [number, number, number, number]
export type HAlign = 'left' | 'center' | 'right' | 'justify'
export type VAlign = 'top' | 'middle' | 'bottom'
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
// used by: image, object, text,
export interface HyperLink {
	slide?: number
	url?: string
	tooltip?: string
}
export interface IHyperLink extends HyperLink {
	rId: number
}
// used by: chart, text
export interface ShadowOptions {
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
export interface ShapeFill {
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
	// FUTURE: beginArrowSize (1-9)
	// FUTURE: endArrowSize (1-9)

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
// used by: chart, slide, table, text
export interface TextOptions {
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
				 * Bullet character code (unicode)
				 * @since 3.3.0
				 * @example '25BA' // 'BLACK RIGHT-POINTING POINTER' (U+25BA)
				 */
				characterCode?: string
				/**
				 * Indentation (space between bullet and text) (points)
				 * @since 3.3.0
				 * @default 27 // DEF_BULLET_MARGIN
				 * @example 10 // Indents text 10 points from bullet
				 */
				indent?: number
				/**
				 * Number type
				 * @since 3.3.0
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
				 * @since 3.3.0
				 * @default 1
				 * @example 10 // numbered bullets start with 10
				 */
				numberStartAt?: number

				// DEPRECATED

				/**
				 * Bullet code (unicode)
				 * @deprecated 3.3.0 - use `characterCode`
				 */
				code?: string
				/**
				 * Margin between bullet and text
				 * @since 3.2.1
				 * @deplrecated 3.3.0 - use `indent`
				 */
				marginPt?: number
				/**
				 * Number to start with (only applies to type:number)
				 * @deprecated 3.3.0 - use `numberStartAt`
				 */
				startAt?: number
				/**
				 * Number type
				 * @deprecated 3.3.0 - use `numberType`
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

// image / media ==================================================================================
export type MediaType = 'audio' | 'online' | 'video'

export interface ImageOpts extends PositionOptions, OptsDataOrPath {
	hyperlink?: IHyperLink
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
export interface IImageOpts extends ImageOpts {
	placeholder?: any
}
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

// shapes =========================================================================================

export interface ShapeOptions extends PositionOptions {
	/**
	 * Horizontal alignment
	 * @default 'left'
	 */
	align?: HAlign
	/**
	 * Shape fill color properties
	 * @example { color:'FF0000' } // hex string (red)
	 * @example { color:'pptx.SchemeColor.accent1' } // theme color Accent1
	 * @example { color:'0088CC', transparency:50 } // 50% transparent color
	 */
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
	/**
	 * Line options
	 */
	line?: ShapeLine
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
	shadow?: ShadowOptions

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
// TODO: WIP: rename to...? `AddTableOptions` ?
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
	rows: ITableRow[]
}
export type TableRow = number[] | string[] | TableCell[] // TODO: 20200523: Consistency: Remove `number[]` as Cell/IText only take strings
// [internal below]
export interface ITableToSlidesOpts extends TableToSlidesOpts {
	_arrObjTabHeadRows?: TableRow[]
	masterSlide?: ISlideLayout
}
export interface ITableCell extends TableCell {
	type: SLIDE_OBJECT_TYPES.tablecell
	lines?: string[]
	lineHeight?: number
	hmerge?: boolean
	vmerge?: boolean
	optImp?: any
}
export interface ITableOptions extends TableOptions {
	_arrObjTabHeadRows?: TableRow[]
}
export type ITableRow = ITableCell[]

// text ===========================================================================================
export interface GlowOptions {
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

// TODO: WIP: rename to...? `AddTextOptions` ?
export interface ITextOpts extends PositionOptions, OptsDataOrPath, TextOptions {
	/**
	 * Whather "Fit to Shape?" is enabled
	 * @default false
	 */
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
	/**
	 * Character spacing
	 */
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
	glow?: GlowOptions
	hyperlink?: IHyperLink
	indentLevel?: number
	inset?: number
	isTextBox?: boolean
	line?: ShapeLine
	lineIdx?: number
	lineSpacing?: number
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
	shadow?: ShadowOptions
	shape?: SHAPE_NAME
	shrinkText?: boolean
	strike?: boolean
	subscript?: boolean
	superscript?: boolean
	underline?: boolean
	valign?: VAlign
	vert?: 'eaVert' | 'horz' | 'mongolianVert' | 'vert' | 'vert270' | 'wordArtVert' | 'wordArtVertRtl'
	wrap?: boolean

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
	 * @deprecated v3.3.0 - use `line.size`
	 */
	lineSize?: number
	/**
	 * @deprecated v3.3.0 - use `arrowTypeEnd`
	 */
	lineTail?: 'none' | 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
}
export interface IText {
	text: string
	options?: ITextOpts
}

// charts =========================================================================================
// FUTURE: BREAKING-CHANGE: (soln: use `OptsDataLabelPosition|string` until 3.5/4.0)
/*
export interface OptsDataLabelPosition {
	pie: 'ctr' | 'inEnd' | 'outEnd' | 'bestFit'
	scatter: 'b' | 'ctr' | 'l' | 'r' | 't'
	// TODO: add all othere chart types
}
*/

export type ChartAxisTickMark = 'none' | 'inside' | 'outside' | 'cross'
export interface OptsChartData {
	index?: number
	labels?: string[]
	name?: string
	sizes?: number[]
	values?: number[]
}
export interface OptsChartGridLine {
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
	/**
	 * Axis position
	 */
	axisPos?: 'b' | 'l' | 'r' | 't'
	border?: BorderOptions
	chartColors?: HexColor[]
	/**
	 * opacity (0.0 - 1.0)
	 * @example 0.5 // 50% opaque
	 */
	chartColorsOpacity?: number
	dataBorder?: BorderOptions
	displayBlanksAs?: string
	fill?: HexColor
	invertedColors?: string
	lang?: string
	layout?: PositionOptions
	shadow?: ShadowOptions
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
	/**
	 * Value label format code
	 * - this also directs Data Table formatting
	 * @since 3.3.0
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
	/**
	 * Data table format code
	 * @since 3.3.0
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
export interface IChartOptsLib extends IChartOpts {
	_type?: CHART_NAME | IChartMulti[]
}

// Core
// ====
/**
 * Section options
 */
export interface ISectionProps {
	title: string
	order?: number // 1-n
}
export interface ISection {
	type: 'user' | 'default'
	title: string
	slides: ISlideLib[]
}
/**
 * The Presentation Layout (ex: 'LAYOUT_WIDE')
 */
export interface ILayout {
	name: string
	width?: number
	height?: number
}
export interface ILayoutProps {
	name: string
	width: number
	height: number
}
export interface ISlideNumber extends PositionOptions, TextOptions {
	align?: HAlign
	color?: string
}
export interface ISlideMasterOptions {
	title: string
	height?: number
	width?: number
	margin?: Margin
	background?: BkgdOpts
	bkgd?: string | BkgdOpts // @deprecated v3.3.0
	objects?: (
		| { chart: {} }
		| { image: {} }
		| { line: {} }
		| { rect: {} }
		| { text: { options: ITextOpts; text?: string } }
		| { placeholder: { options: ISlideMstrObjPlchldrOpts; text?: string } }
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
export interface ISlideRelChart extends OptsChartData {
	type: CHART_NAME | IChartMulti[]
	opts: IChartOptsLib
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
	svgSize?: { w: number; h: number }
	rId: number
	Target: string
}
// TODO: create `ObjectOptions` (placeholder props are internal)
export interface IObjectOptions extends ShapeOptions, TableCellOpts, ITextOpts {
	x?: Coord
	y?: Coord
	cx?: Coord
	cy?: Coord
	w?: Coord
	h?: Coord
	margin?: Margin
	// table
	colW?: number | number[]
	rowH?: number | number[]
	// image:
	sizing?: {
		type?: string
		x?: number
		y?: number
		w?: number
		h?: number
	}
	rounding?: string
	// placeholder
	placeholderIdx?: number
	placeholderType?: PLACEHOLDER_TYPES
}
export interface ISlideObject {
	type: SLIDE_OBJECT_TYPES
	options?: IObjectOptions
	// text
	text?: string | IText[]
	// table
	arrTabRows?: ITableCell[][]
	// chart
	chartRid?: number
	// image:
	image?: string
	imageRid?: number
	hyperlink?: IHyperLink
	// media
	media?: string
	mtype?: MediaType
	mediaRid?: number
	shape?: SHAPE_NAME
}

export interface ISlideLayout {
	presLayout: ILayout
	name: string
	number: number
	background?: BkgdOpts
	bkgd?: string // @deprecated v3.3.0
	bkgdImgRid?: number
	slide?: {
		back: string
		bkgdImgRid?: number
		color: string
		hidden?: boolean
	}
	data: ISlideObject[]
	rels: ISlideRel[]
	relsChart: ISlideRelChart[] // needed as we use args:"ISlide|ISlideLayout" often
	relsMedia: ISlideRelMedia[] // needed as we use args:"ISlide|ISlideLayout" often
	margin?: Margin
	slideNumberObj?: ISlideNumber
}
export interface IAddSlideOptions {
	masterName?: string // TODO: 20200528: rename to "masterTitle" (createMaster uses `title` so lets be consistent)
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
	background?: BkgdOpts
	bkgd?: string // @deprecated v3.3.0
	/**
	 * Default text color (hex format)
	 * @example 'FF3399'
	 */
	color?: HexColor
	hidden?: boolean
	slideNumber?: ISlideNumber
}
export interface ISlideLib extends ISlide {
	bkgdImgRid?: number // FIXME rename
	data?: ISlideObject[]
	id: number
	margin?: Margin
	name?: string
	number: number
	presLayout: ILayout
	rels: ISlideRel[]
	relsChart: ISlideRelChart[]
	relsMedia: ISlideRelMedia[]
	rId: number
	slideLayout: ISlideLayout
	slideNumberObj?: ISlideNumber // FIXME rename
}
export interface IPresentation {
	author: string
	company: string
	layout: string
	masterSlide: ISlide
	presLayout: ILayout
	revision: string
	rtlMode: boolean
	sections: ISection[]
	slideLayouts: ISlideLayout[]
	slides: ISlideLib[]
	subject: string
	title: string
}
