// trigger does not work properly
/**
 * PptxGenJS: Animation Generation
 * Generates animation XML for PowerPoint slides
 * Integrates with gen-xml.ts slide generation
 */

/**
 * Animation preset IDs mapping
 * These correspond to PowerPoint's built-in animation effects
 */
/**
 * Animation preset IDs mapping
 * These correspond to PowerPoint's built-in animation effects
 * Organized by animation class: entrance, emphasis, exit, and path
 */

import type { 
	AnimationConfig, 
	AnimationTrigger, 
	AnimationType, 
	ColorAnimationConfig, 
	FloatAnimationConfig, 
	FlyAnimationConfig,
	GrowShrinkAnimationConfig,
	RandomBarsAnimationConfig,
	ShapeAnimationConfig,
	SlideObjectAnimation,
	SpinAnimationConfig,
	SplitAnimationConfig,
	TransparencyAnimationConfig,
	WheelAnimationConfig,
	WipeAnimationConfig,
	ZoomAnimationConfig
} from './core-interfaces'


export const ANIMATION_PRESETS: Record<string, { 
	id: number
	class: string
	filter?: string
	subtype?: number
	subtypes?: Record<string, number>
}> = {
	// ============================================
	// ENTRANCE ANIMATIONS
	// ============================================
	
	// Basic Entrance
	appear: { 
		id: 1, 
		class: 'entr', 
		subtype: 0 
	},
	
	fadein: { 
		id: 10, 
		class: 'entr', 
		subtype: 0 
	},
	
	// Directional Entrance
	flyin: {
		id: 2,
		class: 'entr',
		subtypes: {
			bottom: 4,
			bottomLeft: 12,
			left: 8,
			topLeft: 9,
			top: 1,
			topRight: 3,
			right: 2,
			bottomRight: 6,
		},
	},
	
	floatin: { 
		id: 42, 
		class: 'entr', 
		subtype: 0 
	},
	
	// Shape-based Entrance
	split: {
		id: 41,
		class: 'entr',
		filter: 'barn',
		subtypes: {
			horizontalIn: 26,
			horizontalOut: 42,
			verticalIn: 21,
			verticalOut: 37,
		},
	},
	
	wipe: {
		id: 22,
		class: 'entr',
		filter: 'wipe',
		subtypes: {
			bottom: 4,
			top: 1,
			left: 8,
			right: 2,
		},
	},
	
	shape: {
		id: 6,
		class: 'entr',
		subtypes: {
			circle: 16,
			diamond: 32,
			plus: 48,
			box: 16,
		},
	},
	
	wheel: {
		id: 21,
		class: 'entr',
		filter: 'wheel',
		subtypes: {
			spoke1: 1,
			spoke2: 2,
			spoke3: 3,
			spoke4: 4,
			spoke8: 8,
		},
	},
	
	randombars: {
		id: 14,
		class: 'entr',
		filter: 'randombar',
		subtypes: {
			horizontal: 10,
			vertical: 5,
		},
	},
	
	// Scale-based Entrance
	zoom: { 
		id: 53, 
		class: 'entr', 
		subtype: 16 
	},
	
	grow: { 
		id: 3, 
		class: 'entr', 
		subtype: 0 
	},
	
	growandturn: { 
		id: 31, 
		class: 'entr', 
		subtype: 0 
	},
	
	// Rotation-based Entrance
	swivel: { 
		id: 45, 
		class: 'entr', 
		subtype: 0 
	},
	
	// Dynamic Entrance
	bounce: { 
		id: 26, 
		class: 'entr', 
		subtype: 0 
	},

	// ============================================
	// EMPHASIS ANIMATIONS
	// ============================================
	
	// Basic Emphasis
	pulse: { 
		id: 26, 
		class: 'emph', 
		subtype: 0 
	},
	
	// Color Emphasis
	colorpulse: { 
		id: 27, 
		class: 'emph', 
		subtype: 0 
	},
	
	desaturate: { 
		id: 25, 
		class: 'emph', 
		subtype: 0 
	},
	
	darken: { 
		id: 24, 
		class: 'emph', 
		subtype: 0 
	},
	
	lighten: { 
		id: 30, 
		class: 'emph', 
		subtype: 0 
	},
	
	objectcolor: { 
		id: 19, 
		class: 'emph', 
		subtype: 0 
	},
	
	complementarycolor: { 
		id: 21, 
		class: 'emph', 
		subtype: 0 
	},
	
	linecolor: { 
		id: 7, 
		class: 'emph', 
		subtype: 2 
	},
	
	fillcolor: { 
		id: 1, 
		class: 'emph', 
		subtype: 2 
	},
	
	// Motion Emphasis
	teeter: { 
		id: 32, 
		class: 'emph', 
		subtype: 0 
	},
	
	spin: { 
		id: 8, 
		class: 'emph', 
		subtype: 0 
	},
	
	// Scale Emphasis
	growshrink: { 
		id: 6, 
		class: 'emph', 
		subtype: 0 
	},
	
	// Visual Emphasis
	transparency: { 
		id: 9, 
		class: 'emph', 
		subtype: 0 
	},

	// ============================================
	// EXIT ANIMATIONS
	// ============================================
	
	// Basic Exit
	disappear: { 
		id: 1, 
		class: 'exit', 
		subtype: 0 
	},
	
	fadeout: { 
		id: 10, 
		class: 'exit', 
		subtype: 0 
	},
	
	// Directional Exit
	flyout: {
		id: 2,
		class: 'exit',
		subtypes: {
			bottom: 4,
			top: 1,
			left: 8,
			right: 2,
			bottomLeft: 12,
			topLeft: 9,
			topRight: 3,
			bottomRight: 6,
		},
	},
	
	floatout: { 
		id: 42, 
		class: 'exit', 
		subtype: 0 
	},
	
	// Shape-based Exit
	splitexit: {
		id: 16,
		class: 'exit',
		filter: 'barn',
		subtypes: {
			horizontalIn: 26,
			horizontalOut: 42,
			verticalIn: 21,
			verticalOut: 37,
		},
	},
	
	wipeexit: {
		id: 22,
		class: 'exit',
		filter: 'wipe',
		subtypes: {
			bottom: 4,
			top: 1,
			left: 8,
			right: 2,
		},
	},
	
	shapeexit: {
		id: 6,
		class: 'exit',
		filter: 'circle',
		subtypes: {
			circle: 32,
			diamond: 16,
			box: 32,
			plus: 32,
		},
	},
	
	wheelexit: {
		id: 21,
		class: 'exit',
		filter: 'wheel',
		subtypes: {
			spoke1: 1,
			spoke2: 2,
			spoke3: 3,
			spoke4: 4,
			spoke8: 8,
		},
	},
	
	randombarsexit: {
		id: 14,
		class: 'exit',
		filter: 'randombar',
		subtypes: {
			horizontal: 10,
			vertical: 5,
		},
	},
	
	// Scale-based Exit
	zoomexit: { 
		id: 53, 
		class: 'exit', 
		subtype: 32 
	},
	
	shrinkandturn: { 
		id: 31, 
		class: 'exit', 
		subtype: 0 
	},
	
	// Rotation-based Exit
	swivelexit: { 
		id: 45, 
		class: 'exit', 
		subtype: 0 
	},
	
	// Dynamic Exit
	bounceexit: { 
		id: 26, 
		class: 'exit', 
		subtype: 0 
	},

	// ============================================
	// MOTION PATH ANIMATIONS
	// ============================================
	
	pathdown: { 
		id: 42, 
		class: 'path', 
		subtype: 0 
	},
	
	patharcdown: { 
		id: 37, 
		class: 'path', 
		subtype: 0 
	},
	
	pathturnright: { 
		id: 50, 
		class: 'path', 
		subtype: 0 
	},
	
	pathcircle: { 
		id: 1, 
		class: 'path', 
		subtype: 0 
	},
	
	pathzigzag: { 
		id: 26, 
		class: 'path', 
		subtype: 0 
	},
}

/**
 * Get animation preset information
 */
function getAnimationPreset(type: AnimationType, direction?: string): { id: number; class: string; subtype: number } {
	
	const normalizedType = type.toLowerCase().replace(/[-_\s]/g, '')
	const preset = ANIMATION_PRESETS[normalizedType] || ANIMATION_PRESETS.appear

	if (preset.subtypes && direction) {
		return {
			id: preset.id,
			class: preset.class,
			subtype: preset.subtypes[direction] || 0,
		}
	}

	return {
		id: preset.id,
		class: preset.class,
		subtype: preset.subtype || 0,
	}
}



/**
 * Generate animation effect XML for a single animation
 * This generates the <p:par> block for one animation effect
 */
function genAnimationEffectXml(animation: AnimationConfig, shapeId: number, nodeId: number): string {
	const preset = getAnimationPreset(animation.type, 'direction' in animation ? animation.direction : undefined)
	const duration = animation.duration || 1000
	const delay = animation.delay || 0
	const animClass =  preset.class
	const presetID = preset.id
	const presetSubtype = preset.subtype
	const nodeType = getNodeType(animation.trigger)
		
	let xml = '<p:par>'
	xml += `<p:cTn id="${nodeId}" presetID="${presetID}" presetClass="${animClass}" presetSubtype="${presetSubtype}" fill="hold" grpId="0" nodeType="${nodeType}">`
	xml += `<p:stCondLst><p:cond delay="${delay}"/></p:stCondLst>`
	xml += '<p:childTnLst>'

	// Generate animation-specific XML based on type
	const animType = animation.type.toLowerCase()

	if (animType === 'appear') {
		// Simple visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
	} else if (animType === 'fadein') {
		// Fade effect
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		xml += `<p:animEffect transition="${animClass === 'exit' ? 'out' : 'in'}" filter="fade">`
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
	}
	else if (animType === 'flyin') {
	// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Direction mapping with proper coordinates
		const directions = {
			bottom: {        // presetSubtype="4"
				startX: '#ppt_x',
				endX: '#ppt_x',
				startY: '1+#ppt_h/2',
				endY: '#ppt_y'
			},
			bottomLeft: {    // presetSubtype="12"
				startX: '0-#ppt_w/2',
				endX: '#ppt_x',
				startY: '1+#ppt_h/2',
				endY: '#ppt_y'
			},
			left: {          // presetSubtype="8"
				startX: '0-#ppt_w/2',
				endX: '#ppt_x',
				startY: '#ppt_y',
				endY: '#ppt_y'
			},
			topLeft: {       // presetSubtype="9"
				startX: '0-#ppt_w/2',
				endX: '#ppt_x',
				startY: '0-#ppt_h/2',
				endY: '#ppt_y'
			},
			top: {           // presetSubtype="1"
				startX: '#ppt_x',
				endX: '#ppt_x',
				startY: '0-#ppt_h/2',
				endY: '#ppt_y'
			},
			topRight: {      // presetSubtype="3"
				startX: '1+#ppt_w/2',
				endX: '#ppt_x',
				startY: '0-#ppt_h/2',
				endY: '#ppt_y'
			},
			right: {         // presetSubtype="2"
				startX: '1+#ppt_w/2',
				endX: '#ppt_x',
				startY: '#ppt_y',
				endY: '#ppt_y'
			},
			bottomRight: {   // presetSubtype="6"
				startX: '1+#ppt_w/2',
				endX: '#ppt_x',
				startY: '1+#ppt_h/2',
				endY: '#ppt_y'
			}
		}
		const flyAnim = animation as FlyAnimationConfig
		const dir = directions[flyAnim.direction || 'bottom']

		if(!dir) {
			throw new Error(`Unknown animation direction: ${flyAnim.direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// X position animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr additive="base">'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += `<p:tav tm="0"><p:val><p:strVal val="${dir.startX}"/></p:val></p:tav>`
		xml += `<p:tav tm="100000"><p:val><p:strVal val="${dir.endX}"/></p:val></p:tav>`
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Y position animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr additive="base">'
		xml += `<p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += `<p:tav tm="0"><p:val><p:strVal val="${dir.startY}"/></p:val></p:tav>`
		xml += `<p:tav tm="100000"><p:val><p:strVal val="${dir.endY}"/></p:val></p:tav>`
		xml += '</p:tavLst>'
		xml += '</p:anim>'
	} else if (animType === 'floatin') {
		// Visibility + fade + position
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		xml += '<p:animEffect transition="in" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// X position (stays same)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Y position - direction based on floatUp or floatDown
		const floatAnim = animation as FloatAnimationConfig
		const direction = floatAnim.direction || 'floatUp'
		const startY = direction === 'floatDown' ? '#ppt_y-.1' : '#ppt_y+.1'
		
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += `<p:tav tm="0"><p:val><p:strVal val="${startY}"/></p:val></p:tav>`
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
	}
	else if (animType === 'split') {
		// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Direction mapping for split animations
		const directions = {
			horizontalIn: 'barn(inHorizontal)',   // presetSubtype="26"
			horizontalOut: 'barn(outHorizontal)', // presetSubtype="42"
			verticalIn: 'barn(inVertical)',       // presetSubtype="21"
			verticalOut: 'barn(outVertical)'      // presetSubtype="37"
		}
		const splitAnim = animation as SplitAnimationConfig
		const direction = splitAnim.direction || 'horizontalIn'
		const filterValue = directions[direction]
		
		if (!filterValue) {
			throw new Error(`Unknown split animation direction: ${direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// Split animation effect
		xml += `<p:animEffect transition="in" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
	}
	else if (animType === 'wipe') {
		// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Direction mapping for wipe animations
		const directions = {
			bottom: 'wipe(down)',  // presetSubtype="4"
			left: 'wipe(left)',    // presetSubtype="8"
			right: 'wipe(right)',  // presetSubtype="2"
			top: 'wipe(up)'        // presetSubtype="1"
		}
		const wipeAnim = animation as WipeAnimationConfig
		const direction = wipeAnim.direction || 'bottom'
		const filterValue = directions[direction]
		
		if (!filterValue) {
			throw new Error(`Unknown wipe animation direction: ${direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// Wipe animation effect
		xml += `<p:animEffect transition="in" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
	}
	else if (animType==='shape') {
		// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Shape types mapping
		const shapes = {
			circle: {
				in: 'circle(in)',    // presetID="6", presetSubtype="16"
				out: 'circle(out)'   // presetID="6", presetSubtype="32"
			},
			box: {
				in: 'box(in)',       // presetID="4", presetSubtype="16"
				out: 'box(out)'      // presetID="4", presetSubtype="32"
			},
			diamond: {
				in: 'diamond(in)',   // presetID="8", presetSubtype="16"
				out: 'diamond(out)'  // presetID="8", presetSubtype="32"
			},
			plus: {
				in: 'plus(in)',      // presetID="13", presetSubtype="16"
				out: 'plus(out)'     // presetID="13", presetSubtype="32"
			}
		}
		const shapeAnim = animation as ShapeAnimationConfig
		const shapeType = shapeAnim.shape || 'circle'
		const direction = shapeAnim.direction || 'in'
		
		if (!shapes[shapeType]) {
			throw new Error(`Unknown shape animation type: ${shapeType}. Valid shapes are: ${Object.keys(shapes).join(', ')}`)
		}
		
		if (!shapes[shapeType][direction]) {
			throw new Error(`Unknown shape animation direction: ${direction}. Valid directions are: in, out`)
		}
		
		const filterValue = shapes[shapeType][direction]
		
		// Shape animation effect
		xml += `<p:animEffect transition="in" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
	}else if (animType==='wheel') {
		// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Wheel spoke count mapping
		// Based on the XML: presetID="21" (wheel animation), presetSubtype indicates spokes
		const spokes = {
			1: 'wheel(1)',  // 1 spoke - presetSubtype="1"
			2: 'wheel(2)',  // 2 spokes - presetSubtype="2"
			3: 'wheel(3)',  // 3 spokes - presetSubtype="3"
			4: 'wheel(4)',  // 4 spokes - presetSubtype="4"
			8: 'wheel(8)'   // 8 spokes - presetSubtype="8"
		}
		const wheelAnim = animation as WheelAnimationConfig
		const spokeCount = wheelAnim.spokes || 1
		
		if (!spokes[spokeCount]) {
			throw new Error(`Unknown wheel animation spoke count: ${spokeCount}. Valid spoke counts are: ${Object.keys(spokes).join(', ')}`)
		}
		
		const filterValue = spokes[spokeCount]
		
		// Wheel animation effect
		xml += `<p:animEffect transition="in" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
	}
	else if (animType === 'randombars'){
		// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Direction mapping for randombar animations
		// Based on the XML: presetID="14" (randombar animation)
		const directions = {
			horizontal: 'randombar(horizontal)',  // presetSubtype="10"
			vertical: 'randombar(vertical)'       // presetSubtype="5"
		}
		const randomBarAnim = animation as RandomBarsAnimationConfig
		const direction = randomBarAnim.direction || 'horizontal'
		const filterValue = directions[direction]
		
		if (!filterValue) {
			throw new Error(`Unknown randombar animation direction: ${direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// RandomBar animation effect
		xml += `<p:animEffect transition="in" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
	}
	else if (animType === 'growandturn') {
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Width animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Height animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Rotation animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.rotation</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:fltVal val="90"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Fade effect
		xml += '<p:animEffect transition="in" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 7}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
	} else if (animType === 'zoom') {

		const zoomAnim = animation as ZoomAnimationConfig
		const isSlideCenter = zoomAnim.direction === 'slideCenter'

		// Visibility set
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Width animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_w"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Height animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Fade effect
		xml += '<p:animEffect transition="in" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// Add position animations ONLY for slideCenter
		if (isSlideCenter) {
			// X position animation
			xml += '<p:anim calcmode="lin" valueType="num">'
			xml += `<p:cBhvr><p:cTn id="${nodeId + 7}" dur="${duration}" fill="hold"/>`
			xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
			xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
			xml += '</p:cBhvr>'
			xml += '<p:tavLst>'
			xml += '<p:tav tm="0"><p:val><p:fltVal val="0.5"/></p:val></p:tav>'
			xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>'
			xml += '</p:tavLst>'
			xml += '</p:anim>'
			
			// Y position animation
			xml += '<p:anim calcmode="lin" valueType="num">'
			xml += `<p:cBhvr><p:cTn id="${nodeId + 8}" dur="${duration}" fill="hold"/>`
			xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
			xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
			xml += '</p:cBhvr>'
			xml += '<p:tavLst>'
			xml += '<p:tav tm="0"><p:val><p:fltVal val="0.5"/></p:val></p:tav>'
			xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_y"/></p:val></p:tav>'
			xml += '</p:tavLst>'
			xml += '</p:anim>'
		}
	}else if (animType === 'swivel') {
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		xml += '<p:animEffect transition="in" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// Width oscillation with sine formula
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0" fmla="#ppt_w*sin(2.5*pi*$)"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="1"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Height stays constant
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_h"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
	} else if (animType === 'bounce') {
		// Complex bounce entrance matching bounce exit structure
		const bounce1Dur = 664  // First bounce duration (33.2% of 2000ms)
		const bounce2Dur = 332  // Second bounce
		const bounce3Dur = 164  // Third bounce
		
		// Initial visibility
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
		
		// Wipe down effect at start
		xml += '<p:animEffect transition="in" filter="wipe(down)">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="580"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// X movement (slide in from left)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="1822" tmFilter="0,0; 0.14,0.36; 0.43,0.73; 0.71,0.91; 1.0,1.0">`
		xml += '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="#ppt_x-0.25"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_x"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Y bounce phases - First fall from above (664ms)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${bounce1Dur}" tmFilter="0.0,0.0;0.25,0.07;0.50,0.2;0.75,0.467;1.0,1.0">`
		xml += '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y-0.333"/></p:val></p:tav>'
		xml += '<p:tav tm="5000"><p:val><p:strVal val="ppt_y-0.303"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.277"/></p:val></p:tav>'
		xml += '<p:tav tm="15000"><p:val><p:strVal val="ppt_y-0.251"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.226"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.177"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.133"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.093"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.059"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.032"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.013"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.003"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// First bounce up (664ms)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 7}" dur="${bounce1Dur}" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">`
		xml += `<p:stCondLst><p:cond delay="${bounce1Dur}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.034"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.065"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.090"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.106"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.111"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.106"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.090"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.065"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.034"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Second bounce (332ms)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 8}" dur="${bounce2Dur}" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">`
		xml += `<p:stCondLst><p:cond delay="${bounce1Dur * 2}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.011"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.022"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.030"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.035"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.037"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.035"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.030"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.022"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.011"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Third bounce (164ms)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 9}" dur="${bounce3Dur}" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5;  0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">`
		xml += '<p:stCondLst><p:cond delay="1656"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.004"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.007"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.010"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.012"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.0123"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.012"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.010"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.007"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.004"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Squash/stretch scale animations (8 animations for bounce effect)
		const scaleDelays = [650, 676, 1312, 1338, 1642, 1668, 1808, 1834]
		const scaleDurations = [26, 166, 26, 166, 26, 166, 26, 166]
		const scaleYValues = ['60000', '100000', '80000', '100000', '90000', '100000', '95000', '100000']
		const decelFlags = [false, true, false, true, false, true, false, true]
		
		for (let i = 0; i < 8; i++) {
			xml += '<p:animScale>'
			xml += `<p:cBhvr><p:cTn id="${nodeId + 10 + i}" dur="${scaleDurations[i]}"${decelFlags[i] ? ' decel="50000"' : ''}>`
			xml += `<p:stCondLst><p:cond delay="${scaleDelays[i]}"/></p:stCondLst></p:cTn>`
			xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
			xml += `<p:to x="100000" y="${scaleYValues[i]}"/>`
			xml += '</p:animScale>'
		}
	}
	else if (animType === 'pulse') {
		// Pulse: fade effect + scale animation
		xml += '<p:animEffect transition="out" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="500" tmFilter="0, 0; .2, .5; .8, .5; 1, 0"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		xml += '<p:animScale>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="250" autoRev="1" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '<p:by x="105000" y="105000"/>'
		xml += '</p:animScale>'
	}
	else if (animType === 'colorpulse') {
		// Default to yellow if no color provided
		const colorAnim = animation as ColorAnimationConfig
		const pulseColor = colorAnim.color || 'FFFF00' // color parameter should be hex without '#'
		
		// Text color animation
		xml += '<p:animClr clrSpc="rgb" dir="cw">'
		xml += '<p:cBhvr override="childStyle">'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}" autoRev="1" fill="remove"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += `<p:to><a:srgbClr val="${pulseColor}"/></p:to>`
		xml += '</p:animClr>'
		
		// Fill color animation
		xml += '<p:animClr clrSpc="rgb" dir="cw">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}" autoRev="1" fill="remove"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += `<p:to><a:srgbClr val="${pulseColor}"/></p:to>`
		xml += '</p:animClr>'
		
		// Set fill type to solid
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 5}" dur="${duration}" autoRev="1" fill="remove"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
		
		// Enable fill
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 6}" dur="${duration}" autoRev="1" fill="remove"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.on</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="true"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'teeter') {
		// Teeter: multiple sequential rotations
		xml += '<p:animRot by="120000">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '</p:animRot>'
		xml += '<p:animRot by="-240000">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="200"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '</p:animRot>'
		xml += '<p:animRot by="240000">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="400"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '</p:animRot>'
		xml += '<p:animRot by="-240000">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="600"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '</p:animRot>'
		xml += '<p:animRot by="120000">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 7}" dur="${duration}" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="800"/></p:stCondLst></p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '</p:animRot>'
	}
	else if (animType === 'spin') {
		// Spin: single rotation
		const SPIN_AMOUNTS = {
			quarterSpin: 5400000,
			halfSpin: 10800000,
			fullSpin: 21600000,
			twoSpins: 43200000,
		}
		const spinAnim = animation as SpinAnimationConfig
		const amount = spinAnim?.amount || 'fullSpin'
		const spinDir = spinAnim?.direction || 'clockwise'

		let rotationValue = SPIN_AMOUNTS[amount]
		if (spinDir === 'counterClockwise') {
			rotationValue = -rotationValue
		}

		xml += `<p:animRot by="${rotationValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '</p:animRot>'
	}
	else if (animType === 'growshrink') {

		const GROW_SHRINK_CONFIG = {
			// Direction determines which axis changes
			direction: {
				horizontal: { x: true, y: false },  // Only X changes
				vertical: { x: false, y: true },    // Only Y changes
				both: { x: true, y: true }          // Both change
			},
			amount: {
				tiny: 50000,      // 50% (shrinks to half)
				smaller: 75000,   // 75%
				larger: 150000,   // 150% (grows 1.5x)
				huge: 400000      // 400% (grows 4x)
			}
		}

		const growShrinkAnim = animation as GrowShrinkAnimationConfig

		const direction = growShrinkAnim?.direction || 'both'
		const amount = growShrinkAnim?.amount || 'larger'
		
		const dirConfig = GROW_SHRINK_CONFIG.direction[direction]
		const scaleValue = GROW_SHRINK_CONFIG.amount[amount]
		
		// Calculate X and Y scale values
		// If direction affects the axis, use the scale value, otherwise use 100000 (100% = no change)
		const scaleX = dirConfig.x ? scaleValue : 100000
		const scaleY = dirConfig.y ? scaleValue : 100000
		
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += `<p:by x="${scaleX}" y="${scaleY}"/>`
		xml += '</p:animScale>'
	}
	else if (animType === 'desaturate') {
		// Desaturate: reduce saturation
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr override="childStyle"><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="-70588" l="0"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="-70588" l="0"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>stroke.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="-70588" l="0"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'darken') {
		// Darken: reduce lightness
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr override="childStyle"><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="-12549" l="-25098"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="-12549" l="-25098"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>stroke.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="-12549" l="-25098"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'lighten') {
		// Lighten: increase lightness
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr override="childStyle"><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="12549" l="25098"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="500" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="12549" l="25098"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="500" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>stroke.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="0" s="12549" l="25098"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="500" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'transparency') {

		// Transparency
		const TRANSPARENCY_LEVELS = {
			'25%': '0.75',   // 25% transparent = 75% opaque
			'50%':'0.5',    // 50% transparent = 50% opaque
			'75%': '0.2',   // 75% transparent = 25% opaque
			'100%': '0',     // 100% transparent = fully transparent
		}


		let opacityValue = '0.5'
		const transparencyAnim  = animation as TransparencyAnimationConfig

		opacityValue = transparencyAnim?.level ? TRANSPARENCY_LEVELS[transparencyAnim.level] : '0.5'
		
		// Set opacity
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += `<p:to><p:strVal val="${opacityValue}"/></p:to>`
		xml += '</p:set>'
		
		// Image filter effect (for compatibility)
		xml += `<p:animEffect filter="image" prLst="opacity: ${opacityValue}">`
		xml += '<p:cBhvr rctx="IE">'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
	}
	else if (animType === 'objectcolor') {

		const colorAnim = animation as ColorAnimationConfig
		// Object Color: change fill and text color
		xml += '<p:animClr clrSpc="rgb" dir="cw">'
		xml += `<p:cBhvr override="childStyle"><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += `<p:to><a:srgbClr val="${colorAnim.color || 'FFC000'}"/></p:to>`
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="rgb" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += `<p:to><a:srgbClr val="${colorAnim.color || 'FFC000'}"/></p:to>`
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.on</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="true"/></p:to>'
		xml += '</p:set>'
	}else if (animType === 'complementarycolor') {
		// Complementary Color: rotate hue by 180 degrees (7200000 = 180 degrees in PowerPoint units)
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr override="childStyle"><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="7200000" s="0" l="0"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="7200000" s="0" l="0"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:animClr clrSpc="hsl" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>stroke.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:by><p:hsl h="7200000" s="0" l="0"/></p:by>'
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
	}else if (animType === 'linecolor') {
		const colorAnim = animation as ColorAnimationConfig
		// Line Color: change stroke color
		xml += '<p:animClr clrSpc="rgb" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>stroke.color</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += `<p:to><a:srgbClr val="${colorAnim.color || 'E7E6E6'}"/></p:to>`
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>stroke.on</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="true"/></p:to>'
		xml += '</p:set>'
	}else if (animType === 'fillcolor') {
		const colorAnim = animation as ColorAnimationConfig
		// Fill Color: change fill color
		xml += '<p:animClr clrSpc="rgb" dir="cw">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fillcolor</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += `<p:to><a:srgbClr val="${colorAnim.color || 'FFC000'}"/></p:to>`
		xml += '</p:animClr>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.type</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="solid"/></p:to>'
		xml += '</p:set>'
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}" fill="hold"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>fill.on</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="true"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType.includes('path')) {
		let path = 'M 0 0 L 0 0.25 E'
		if (animType === 'pathcircle') path = 'M 0 0 C 0.069 0 0.125 0.056 0.125 0.125 C 0.125 0.194 0.069 0.25 0 0.25 C -0.069 0.25 -0.125 0.194 -0.125 0.125 C -0.125 0.056 -0.069 0 0 0 Z'
		else if (animType === 'patharcdown') path = 'M 0 0 L 0.067 0.04 C 0.081 0.049 0.102 0.054 0.124 0.054 C 0.149 0.054 0.169 0.049 0.183 0.04 L 0.25 0 E'
		else if (animType === 'pathturnright') path = 'M 0 0 L 0.125 0 C 0.181 0 0.25 0.069 0.25 0.125 L 0.25 0.25 E'
		else if (animType === 'pathzigzag') path = 'M 0 0 C 0 0.033 0.027 0.06 0.06 0.06 C 0.099 0.06 0.113 0.03 0.119 0.012 L 0.125 -0.012 C 0.131 -0.03 0.146 -0.06 0.19 -0.06 C 0.218 -0.06 0.25 -0.033 0.25 0 C 0.25 0.033 0.218 0.06 0.19 0.06 C 0.146 0.06 0.131 0.03 0.125 0.012 L 0.119 -0.012 C 0.113 -0.03 0.099 -0.06 0.06 -0.06 C 0.027 -0.06 0 -0.033 0 0 Z'

		xml += `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" ptsTypes="">`
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}" fill="hold"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '</p:animMotion>'
	} 

	else if (animType === 'disappear') {
		// Disappear: hide the element instantly
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'fadeout') {
		// Fade out effect
		xml += '<p:animEffect transition="out" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// Set visibility to hidden at the end (delay = duration - 1)
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'flyout') {
		// Direction mapping with proper coordinates for fly-out
		const directions = {
			bottom: {           // presetSubtype="4"
				startX: 'ppt_x',
				endX: 'ppt_x',
				startY: 'ppt_y',
				endY: '1+ppt_h/2'
			},
			bottomLeft: {       // presetSubtype="12"
				startX: 'ppt_x',
				endX: '0-ppt_w/2',
				startY: 'ppt_y',
				endY: '1+ppt_h/2'
			},
			left: {             // presetSubtype="8"
				startX: 'ppt_x',
				endX: '0-ppt_w/2',
				startY: 'ppt_y',
				endY: 'ppt_y'
			},
			topLeft: {          // presetSubtype="9"
				startX: 'ppt_x',
				endX: '0-ppt_w/2',
				startY: 'ppt_y',
				endY: '0-ppt_h/2'
			},
			top: {              // presetSubtype="1"
				startX: 'ppt_x',
				endX: 'ppt_x',
				startY: 'ppt_y',
				endY: '0-ppt_h/2'
			},
			topRight: {         // presetSubtype="3"
				startX: 'ppt_x',
				endX: '1+ppt_w/2',
				startY: 'ppt_y',
				endY: '0-ppt_h/2'
			},
			right: {            // presetSubtype="2"
				startX: 'ppt_x',
				endX: '1+ppt_w/2',
				startY: 'ppt_y',
				endY: 'ppt_y'
			},
			bottomRight: {      // presetSubtype="6"
				startX: 'ppt_x',
				endX: '1+ppt_w/2',
				startY: 'ppt_y',
				endY: '1+ppt_h/2'
			}
		}
		const flyAnim = animation as FlyAnimationConfig
		const dir = directions[flyAnim.direction || 'bottom']
		
		// X position animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr additive="base">'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += `<p:tav tm="0"><p:val><p:strVal val="${dir.startX}"/></p:val></p:tav>`
		xml += `<p:tav tm="100000"><p:val><p:strVal val="${dir.endX}"/></p:val></p:tav>`
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Y position animation
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr additive="base">'
		xml += `<p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += `<p:tav tm="0"><p:val><p:strVal val="${dir.startY}"/></p:val></p:tav>`
		xml += `<p:tav tm="100000"><p:val><p:strVal val="${dir.endY}"/></p:val></p:tav>`
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Visibility set to hidden at the end
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 5}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	} else if (animType === 'floatout') {
	// Float Out: fade out + position change + hide
		xml += '<p:animEffect transition="out" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// X position (stays same)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_x"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_x"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		const floatAnim = animation as FloatAnimationConfig
		// Y position - direction based on floatUp or floatDown
		const direction = floatAnim.direction || 'floatUp'
		const endY = direction === 'floatDown' ? 'ppt_y+.1'  :  'ppt_y-.1'
		
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += `<p:tav tm="100000"><p:val><p:strVal val="${endY}"/></p:val></p:tav>`
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Hide at the end
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'splitexit') {
		// Direction mapping for split exit animations
		const directions = {
			horizontalIn: 'barn(inHorizontal)',   // presetSubtype="26"
			horizontalOut: 'barn(outHorizontal)', // presetSubtype="42"
			verticalIn: 'barn(inVertical)',       // presetSubtype="21"
			verticalOut: 'barn(outVertical)'      // presetSubtype="37"
		}
		const splitAnim = animation as SplitAnimationConfig
		const direction = splitAnim.direction || 'horizontalIn'
		const filterValue = directions[direction]
		
		if (!filterValue) {
			throw new Error(`Unknown split exit direction: ${direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// Split exit animation effect
		xml += `<p:animEffect transition="out" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
		
		// Hide at the end
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'wipeexit') {
	// Direction mapping for wipe exit animations
		const directions = {
			bottom: 'wipe(down)',  // presetSubtype="4"
			left: 'wipe(left)',    // presetSubtype="8"
			right: 'wipe(right)',  // presetSubtype="2"
			top: 'wipe(up)'        // presetSubtype="1"
		}
		const wipeAnim = animation as WipeAnimationConfig
		const direction = wipeAnim.direction || 'bottom'
		const filterValue = directions[direction]
		
		if (!filterValue) {
			throw new Error(`Unknown wipe exit direction: ${direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// Wipe exit animation effect
		xml += `<p:animEffect transition="out" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
		
		// Hide at the end
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'shapeexit') {
		// Shape types mapping
		const shapes = {
			circle: {
				in: 'circle(in)',    // presetID="6", presetSubtype="16"
				out: 'circle(out)'   // presetID="6", presetSubtype="32"
			},
			box: {
				in: 'box(in)',       // presetID="4", presetSubtype="16"
				out: 'box(out)'      // presetID="4", presetSubtype="32"
			},
			diamond: {
				in: 'diamond(in)',   // presetID="8", presetSubtype="16"
				out: 'diamond(out)'  // presetID="8", presetSubtype="32"
			},
			plus: {
				in: 'plus(in)',      // presetID="13", presetSubtype="16"
				out: 'plus(out)'     // presetID="13", presetSubtype="32"
			}
		}
		const shapeAnim = animation as ShapeAnimationConfig
		const shapeType = shapeAnim.shape || 'circle'
		const direction = shapeAnim.direction || 'in'
		
		if (!shapes[shapeType]) {
			throw new Error(`Unknown shape exit type: ${shapeType}. Valid shapes are: ${Object.keys(shapes).join(', ')}`)
		}
		
		if (!shapes[shapeType][direction]) {
			throw new Error(`Unknown shape exit direction: ${direction}. Valid directions are: in, out`)
		}
		
		const filterValue = shapes[shapeType][direction]
		
		// Shape exit animation effect
		xml += `<p:animEffect transition="out" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
		
		// Hide at the end
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'wheelexit') {
		// Wheel animation effect (exit)
		const spokes = {
			1: 'wheel(1)',  // 1 spoke - presetSubtype="1"
			2: 'wheel(2)',  // 2 spokes - presetSubtype="2"
			3: 'wheel(3)',  // 3 spokes - presetSubtype="3"
			4: 'wheel(4)',  // 4 spokes - presetSubtype="4"
			8: 'wheel(8)'   // 8 spokes - presetSubtype="8"
		}
		const wheelAnim = animation as WheelAnimationConfig
		const spokeCount = wheelAnim.spokes || 1
		
		if (!spokes[spokeCount]) {
			throw new Error(`Unknown wheel animation spoke count: ${spokeCount}. Valid spoke counts are: ${Object.keys(spokes).join(', ')}`)
		}
		
		const filterValue = spokes[spokeCount]
		
		xml += `<p:animEffect transition="out" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
		
		// Visibility set to hidden after animation
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'randombarsexit') {
		// Direction mapping for randombar exit animations
		const directions = {
			horizontal: 'randombar(horizontal)',  // presetSubtype="10"
			vertical: 'randombar(vertical)'       // presetSubtype="5"
		}
		const randomBarsAnim = animation as RandomBarsAnimationConfig
		const direction = randomBarsAnim.direction || 'horizontal'
		const filterValue = directions[direction]
		
		if (!filterValue) {
			throw new Error(`Unknown randombar animation direction: ${direction}. Valid directions are: ${Object.keys(directions).join(', ')}`)
		}
		
		// RandomBar exit animation effect
		xml += `<p:animEffect transition="out" filter="${filterValue}">`
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
		
		// Visibility set - hide at the end
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst>`
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'shrinkandturn') {
		// Shrink and turn exit - shrink width and height while rotating
		
		// Width animation - shrink to 0
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Height animation - shrink to 0
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_h"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Rotation animation - rotate 90 degrees
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.rotation</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="90"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Fade out effect
		xml += '<p:animEffect transition="out" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// Set visibility to hidden at the end
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 7}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'zoomexit') {
		const zoomAnim = animation as ZoomAnimationConfig
		const isSlideCenter = zoomAnim.direction === 'slideCenter'
		
		// Width animation - shrink to 0
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 1}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Height animation - shrink to 0
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 2}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_h"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Fade effect (out)
		xml += '<p:animEffect transition="out" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// Add position animations ONLY for slideCenter
		if (isSlideCenter) {
			// X position animation - move to center (0.5)
			xml += '<p:anim calcmode="lin" valueType="num">'
			xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/>`
			xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
			xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
			xml += '</p:cBhvr>'
			xml += '<p:tavLst>'
			xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_x"/></p:val></p:tav>'
			xml += '<p:tav tm="100000"><p:val><p:fltVal val="0.5"/></p:val></p:tav>'
			xml += '</p:tavLst>'
			xml += '</p:anim>'
			
			// Y position animation - move to center (0.5)
			xml += '<p:anim calcmode="lin" valueType="num">'
			xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}"/>`
			xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
			xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
			xml += '</p:cBhvr>'
			xml += '<p:tavLst>'
			xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
			xml += '<p:tav tm="100000"><p:val><p:fltVal val="0.5"/></p:val></p:tav>'
			xml += '</p:tavLst>'
			xml += '</p:anim>'
		}
		
		// Visibility set - hide at the end
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + (isSlideCenter ? 6 : 4)}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst>`
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'swivelexit') {
		// Swivel exit - fade with width oscillation
		xml += '<p:animEffect transition="out" filter="fade">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 3}" dur="${duration}"/><p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl></p:cBhvr>`
		xml += '</p:animEffect>'
		
		// Width animation - creates the swivel effect
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 4}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_w</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="5000"><p:val><p:strVal val="0.92*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="0.71*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="15000"><p:val><p:strVal val="0.38*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="25000"><p:val><p:strVal val="-0.38*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="-0.71*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="35000"><p:val><p:strVal val="-0.92*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="-ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="45000"><p:val><p:strVal val="-0.92*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="-0.71*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="55000"><p:val><p:strVal val="-0.38*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '<p:tav tm="65000"><p:val><p:strVal val="0.38*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="0.71*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="75000"><p:val><p:strVal val="0.92*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="85000"><p:val><p:strVal val="0.92*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="0.71*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="95000"><p:val><p:strVal val="0.38*ppt_w"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:fltVal val="0"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Height animation - stays constant
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 5}" dur="${duration}"/>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_h</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_h"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_h"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Set visibility to hidden at the end
		xml += '<p:set>'
		xml += `<p:cBhvr><p:cTn id="${nodeId + 6}" dur="1" fill="hold">`
		xml += `<p:stCondLst><p:cond delay="${duration - 1}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	}
	else if (animType === 'bounceexit' ) {
		// This is a complex bounce exit animation with multiple stages
		// Total duration appears to be 2000ms in the XML
		
		// Main wipe effect with delay
		xml += '<p:animEffect transition="out" filter="wipe(down)">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 1}" dur="180" accel="50000">`
		xml += '<p:stCondLst><p:cond delay="1820"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '</p:animEffect>'
		
		// X position animation - horizontal drift
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 2}" dur="1822" tmFilter="0,0; 0.14,0.31; 0.43,0.73; 0.71,0.91; 1.0,1.0">`
		xml += '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_x"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="#ppt_x+0.25"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// X position hold
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="178">`
		xml += '<p:stCondLst><p:cond delay="1822"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_x"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_x"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// First bounce - Y position (main bounce down)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 4}" dur="664" tmFilter="0.0,0.0;0.25,0.07;0.50,0.2;0.75,0.467;1.0,1.0">`
		xml += '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="5000"><p:val><p:strVal val="ppt_y+0.026"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y+0.052"/></p:val></p:tav>'
		xml += '<p:tav tm="15000"><p:val><p:strVal val="ppt_y+0.078"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y+0.103"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y+0.151"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y+0.196"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y+0.236"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y+0.270"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y+0.297"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y+0.317"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y+0.329"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y+0.333"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Second bounce - Y position (bounce up and return)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 5}" dur="664" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5; 0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">`
		xml += '<p:stCondLst><p:cond delay="664"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.034"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.065"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.090"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.106"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.111"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.106"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.090"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.065"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.034"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Third bounce - Y position (smaller bounce)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 6}" dur="332" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5; 0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">`
		xml += '<p:stCondLst><p:cond delay="1324"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.011"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.022"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.030"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.035"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.037"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.035"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.030"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.022"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.011"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Fourth bounce - Y position (even smaller bounce)
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 7}" dur="164" tmFilter="0, 0; 0.125,0.2665; 0.25,0.4; 0.375,0.465; 0.5,0.5; 0.625,0.535; 0.75,0.6; 0.875,0.7335; 1,1">`
		xml += '<p:stCondLst><p:cond delay="1656"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="10000"><p:val><p:strVal val="ppt_y-0.004"/></p:val></p:tav>'
		xml += '<p:tav tm="20000"><p:val><p:strVal val="ppt_y-0.007"/></p:val></p:tav>'
		xml += '<p:tav tm="30000"><p:val><p:strVal val="ppt_y-0.010"/></p:val></p:tav>'
		xml += '<p:tav tm="40000"><p:val><p:strVal val="ppt_y-0.012"/></p:val></p:tav>'
		xml += '<p:tav tm="50000"><p:val><p:strVal val="ppt_y-0.0123"/></p:val></p:tav>'
		xml += '<p:tav tm="60000"><p:val><p:strVal val="ppt_y-0.012"/></p:val></p:tav>'
		xml += '<p:tav tm="70000"><p:val><p:strVal val="ppt_y-0.010"/></p:val></p:tav>'
		xml += '<p:tav tm="80000"><p:val><p:strVal val="ppt_y-0.007"/></p:val></p:tav>'
		xml += '<p:tav tm="90000"><p:val><p:strVal val="ppt_y-0.004"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Final Y drop off screen
		xml += '<p:anim calcmode="lin" valueType="num">'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 8}" dur="180" accel="50000">`
		xml += '<p:stCondLst><p:cond delay="1820"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>ppt_y</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:tavLst>'
		xml += '<p:tav tm="0"><p:val><p:strVal val="ppt_y"/></p:val></p:tav>'
		xml += '<p:tav tm="100000"><p:val><p:strVal val="ppt_y+ppt_h"/></p:val></p:tav>'
		xml += '</p:tavLst>'
		xml += '</p:anim>'
		
		// Scale animations for squash/stretch effect
		// First squash
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 9}" dur="26">`
		xml += '<p:stCondLst><p:cond delay="620"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="60000"/>'
		xml += '</p:animScale>'
		
		// First stretch back
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 10}" dur="166" decel="50000">`
		xml += '<p:stCondLst><p:cond delay="646"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="100000"/>'
		xml += '</p:animScale>'
		
		// Second squash
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 11}" dur="26">`
		xml += '<p:stCondLst><p:cond delay="1312"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="80000"/>'
		xml += '</p:animScale>'
		
		// Second stretch back
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 12}" dur="166" decel="50000">`
		xml += '<p:stCondLst><p:cond delay="1338"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="100000"/>'
		xml += '</p:animScale>'
		
		// Third squash
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 13}" dur="26">`
		xml += '<p:stCondLst><p:cond delay="1642"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="90000"/>'
		xml += '</p:animScale>'
		
		// Third stretch back
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 14}" dur="166" decel="50000">`
		xml += '<p:stCondLst><p:cond delay="1668"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="100000"/>'
		xml += '</p:animScale>'
		
		// Fourth squash
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 15}" dur="26">`
		xml += '<p:stCondLst><p:cond delay="1808"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="95000"/>'
		xml += '</p:animScale>'
		
		// Fourth stretch back
		xml += '<p:animScale>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 16}" dur="166" decel="50000">`
		xml += '<p:stCondLst><p:cond delay="1834"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '</p:cBhvr>'
		xml += '<p:to x="100000" y="100000"/>'
		xml += '</p:animScale>'
		
		// Visibility set - hide at the end
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 17}" dur="1" fill="hold">`
		xml += '<p:stCondLst><p:cond delay="1999"/></p:stCondLst>'
		xml += '</p:cTn>'
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="hidden"/></p:to>'
		xml += '</p:set>'
	} else {
		// appear animation - simply set visibility to visible at start
		xml += '<p:set>'
		xml += '<p:cBhvr>'
		xml += `<p:cTn id="${nodeId + 3}" dur="1" fill="hold"><p:stCondLst><p:cond delay="${duration}"/></p:stCondLst></p:cTn>`
		xml += `<p:tgtEl><p:spTgt spid="${shapeId}"/></p:tgtEl>`
		xml += '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
		xml += '</p:cBhvr>'
		xml += '<p:to><p:strVal val="visible"/></p:to>'
		xml += '</p:set>'
	}

	xml += '</p:childTnLst></p:cTn></p:par>'

	return xml
}

/**
 * Generate build list XML for animations
 * This tells PowerPoint which shapes have animations
 */
function genBuildListXml(animations: SlideObjectAnimation[]): string {
	let xml = '<p:bldLst>'
	animations.forEach((anim) => {
		const shapeId = anim.objectIndex + 2 // Shapes start at ID 2
		xml += `<p:bldP spid="${shapeId}" grpId="0" animBg="1"/>`
	})
	xml += '</p:bldLst>'
	return xml
}

/**
 * Group animations by trigger type for proper XML structure
 */
interface AnimationGroup {
	onClick: SlideObjectAnimation[]
	withPrevious: SlideObjectAnimation[]
	afterPrevious: SlideObjectAnimation[]
	previousDuration: number
}

function groupAnimationsByTrigger(animations: SlideObjectAnimation[]): AnimationGroup[] {
	const groups: AnimationGroup[] = []
	let currentGroup: AnimationGroup | null = null

	animations.forEach((anim) => {
		const trigger = anim.animation.trigger || 'onClick'

		if (trigger === 'onClick') {
			// Start a new group
			currentGroup = {
				onClick: [anim],
				withPrevious: [],
				afterPrevious: [],
				previousDuration: anim.animation.duration || 1000
			}
			groups.push(currentGroup)
		} else if (trigger === 'withPrevious' && currentGroup) {
			// Add to current group's withPrevious
			currentGroup.withPrevious.push(anim)
		} else if (trigger === 'afterPrevious' && currentGroup) {
			// Add to current group's afterPrevious
			currentGroup.afterPrevious.push(anim)
			// Update duration if this animation is longer
			const totalDuration = (anim.animation.duration || 1000) + (anim.animation.delay || 0)
			currentGroup.previousDuration = Math.max(currentGroup.previousDuration, totalDuration)
		} else {
			// No current group, treat as onClick
			currentGroup = {
				onClick: [anim],
				withPrevious: [],
				afterPrevious: [],
				previousDuration: anim.animation.duration || 1000
			}
			groups.push(currentGroup)
		}
	})

	return groups
}

function getNodeType(trigger?: AnimationTrigger): string {
	switch (trigger) {
		case 'withPrevious':
			return 'withEffect'
		case 'afterPrevious':
			return 'afterEffect'
		case 'onClick':
		default:
			return 'clickEffect'
	}
}

/**
 * Generate complete timing XML for all animations on a slide
 * Called by makeXmlSlide() in gen-xml.ts
 * 
 * @param {SlideObjectAnimation[]} animations - array of shape animations
 * @returns complete timing XML block to insert into slide XML
 */

export function createTimingXml(animations: SlideObjectAnimation[]): string {
	if (!animations || animations.length === 0) {
		return ''
	}

	// Group animations by trigger type
	const groups = groupAnimationsByTrigger(animations)

	let xml = '<p:timing>'
	xml += '<p:tnLst>'
	xml += '<p:par>'
	xml += '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
	xml += '<p:childTnLst>'
	xml += '<p:seq concurrent="1" nextAc="seek">'
	xml += '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
	xml += '<p:childTnLst>'

	let nodeId = 3

	// Process each group
	groups.forEach((group) => {
		// For onClick animations, we need the outer <p:par> wrapper
		if (group.onClick.length > 0 || group.withPrevious.length > 0) {
			xml += '<p:par>'
			xml += `<p:cTn id="${nodeId}" fill="hold">`
			xml += '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
			xml += '<p:childTnLst>'
			nodeId++

			// Inner wrapper for all animations in this click
			xml += '<p:par>'
			xml += `<p:cTn id="${nodeId}" fill="hold">`
			xml += '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
			xml += '<p:childTnLst>'
			nodeId++

			// Generate onClick animation
			group.onClick.forEach((shapeAnim) => {
				const shapeId = shapeAnim.objectIndex + 2
				xml += genAnimationEffectXml(shapeAnim.animation, shapeId, nodeId)
				nodeId += 10
			})

			// Generate withPrevious animations (same level as onClick)
			group.withPrevious.forEach((shapeAnim) => {
				const shapeId = shapeAnim.objectIndex + 2
				xml += genAnimationEffectXml(shapeAnim.animation, shapeId, nodeId)
				nodeId += 10
			})

			xml += '</p:childTnLst>'
			xml += '</p:cTn>'
			xml += '</p:par>'
			xml += '</p:childTnLst>'
			xml += '</p:cTn>'
			xml += '</p:par>'
		}

		// For afterPrevious animations, create separate wrapper with delay
		if (group.afterPrevious.length > 0) {
			xml += '<p:par>'
			xml += `<p:cTn id="${nodeId}" fill="hold">`
			xml += `<p:stCondLst><p:cond delay="${group.previousDuration}"/></p:stCondLst>`
			xml += '<p:childTnLst>'
			nodeId++

			group.afterPrevious.forEach((shapeAnim) => {
				const shapeId = shapeAnim.objectIndex + 2
				xml += genAnimationEffectXml(shapeAnim.animation, shapeId, nodeId)
				nodeId += 10
			})

			xml += '</p:childTnLst>'
			xml += '</p:cTn>'
			xml += '</p:par>'
		}
	})

	xml += '</p:childTnLst>'
	xml += '</p:cTn>'

	// Add previous/next conditions
	xml += '<p:prevCondLst>'
	xml += '<p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
	xml += '</p:prevCondLst>'
	xml += '<p:nextCondLst>'
	xml += '<p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
	xml += '</p:nextCondLst>'

	xml += '</p:seq>'
	xml += '</p:childTnLst>'
	xml += '</p:cTn>'
	xml += '</p:par>'
	xml += '</p:tnLst>'

	// Add build list
	xml += genBuildListXml(animations)

	xml += '</p:timing>'

	return xml
}