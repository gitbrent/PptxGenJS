/**
 * PptxGenJS Interfaces
 */

import { CHART_NAME, PLACEHOLDER_TYPE, SHAPE_NAME, SLIDE_OBJECT_TYPES, TEXT_HALIGN, TEXT_VALIGN, WRITE_OUTPUT_TYPE } from './core-enums'

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
	_rId: number
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
	/**
	 * Alt Text value ("How would you describe this object and its contents to someone who is blind?")
	 * - PowerPoint: [right-click on an image] > "Edit Alt Text..."
	 */
	altText?: string
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
	hyperlink?: HyperlinkProps
	/**
	 * Placeholder type
	 * - values: 'body' | 'header' | 'footer' | 'title' | et. al.
	 * @example 'body'
	 * @see https://docs.microsoft.com/en-us/office/vba/api/powerpoint.ppplaceholdertype
	 */
	placeholder?: string
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
	_arrObjTabHeadRows?: TableRow[]
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
	_arrObjTabHeadRows?: TableRow[]

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
	_type: SLIDE_OBJECT_TYPES.tablecell
	_lines?: string[]
	_lineHeight?: number
	_hmerge?: boolean
	_vmerge?: boolean
	_rowContinue?: number
	_optImp?: any

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
	_bodyProp?: {
		// Note: Many of these duplicated as user options are transformed to _bodyProp options for XML processing
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
	_lineIdx?: number

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
// TODO: 202008: chart types remain with predicated with "I" in v3.3.0 (ran out of time!)
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
export interface IChartOptsLib extends IChartOpts {
	_type?: CHART_NAME | IChartMulti[] // TODO: v3.4.0 - move to `IChartOpts`, remove `IChartOptsLib`
}
export interface ISlideRelChart extends OptsChartData {
	type: CHART_NAME | IChartMulti[]
	opts: IChartOptsLib
	data: OptsChartData[]
	// internal below
	rId: number
	Target: string
	globalId: number
	fileName: string
}

// Core
// ====
// PRIVATE vvv
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
	opts?: MediaProps
	path?: string
	extn?: string
	data?: string | ArrayBuffer
	isSvgPng?: boolean
	svgSize?: { w: number; h: number }
	rId: number
	Target: string
}
export interface ISlideObject {
	_type: SLIDE_OBJECT_TYPES
	options?: ObjectOptions
	// text
	text?: TextProps[]
	// table
	arrTabRows?: TableCell[][]
	// chart
	chartRid?: number
	// image:
	image?: string
	imageRid?: number
	hyperlink?: HyperlinkProps
	// media
	media?: string
	mtype?: MediaType
	mediaRid?: number
	shape?: SHAPE_NAME
}
// PRIVATE ^^^

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
	_type: 'user' | 'default'
	_slides: PresSlide[]

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
	_sizeW?: number
	_sizeH?: number

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
	_placeholderIdx?: number
	_placeholderType?: PLACEHOLDER_TYPE

	cx?: Coord
	cy?: Coord
	margin?: Margin
	colW?: number | number[] // table
	rowH?: number | number[] // table
}
export interface SlideBaseProps {
	_bkgdImgRid?: number
	_margin?: Margin
	_name?: string
	_presLayout: PresLayout
	_rels: ISlideRel[]
	_relsChart: ISlideRelChart[] // needed as we use args:"PresSlide|SlideLayout" often
	_relsMedia: ISlideRelMedia[] // needed as we use args:"PresSlide|SlideLayout" often
	_slideNum: number
	_slideNumberProps?: SlideNumberProps
	_slideObjects?: ISlideObject[]

	background?: BackgroundProps
	/**
	 * @deprecated v3.3.0 - use `background`
	 */
	bkgd?: string | BackgroundProps
}
export interface SlideLayout extends SlideBaseProps {
	_slide?: {
		_bkgdImgRid?: number
		back: string
		color: string
		hidden?: boolean
	}
}
export interface PresSlide extends SlideBaseProps {
	_rId: number
	_slideLayout: SlideLayout
	_slideId: number

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
// PRIVATE interface
export interface IPresentationProps extends PresentationProps {
	sections: SectionProps[]
	slideLayouts: SlideLayout[]
	slides: PresSlide[]
}
