import { DEF_SLIDE_MARGIN_IN, DEF_PRES_LAYOUT_NAME, MASTER_OBJECTS, BASE_SHAPES } from './core-enums'

import Relations from './relations'

import TextElement from './elements/text'
import ShapeElement from './elements/simple-shape'
import PlaceholderTextElement from './elements/placeholder-text'
import ImageElement from './elements/image'
import ChartElement from './elements/chart'
import SlideNumberElement from './elements/slide-number'

export class Master {
	name

	number
	slide = null
	data = []
	margin
	relations

	presLayout

	bkgd
	bkgdImgRid

	placeholders: Map<string, PlaceholderTextElement>

	constructor(title, number, layout) {
		if (!title) throw Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		this.name = title
		this.relations = new Relations()
		this.placeholders = new Map()

		this.presLayout = layout
		this.number = number

		this.margin = DEF_SLIDE_MARGIN_IN
	}

	public get rels() {
		return this.relations.rels
	}

	public get relsChart() {
		return this.relations.relsChart
	}

	public get relsMedia() {
		return this.relations.relsMedia
	}

	configureBackground(bkg) {
		if (typeof bkg === 'object' && (bkg.src || bkg.path || bkg.data)) {
			// Allow the use of only the data key (`path` isnt reqd)
			bkg.src = bkg.src || bkg.path || null
			if (!bkg.src) bkg.src = 'preencoded.png'

			// Handle "blah.jpg?width=540" etc.
			let strImgExtn = (bkg.src.split('.').pop() || 'png').split('?')[0]
			// base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]",
			// so correct exttnesion to avoid content warnings at PPT startup
			if (strImgExtn === 'jpg') strImgExtn = 'jpeg'
			this.bkgdImgRid = this.relations.registerImage({ data: bkg.data, path: bkg.src }, strImgExtn)
		} else if (bkg && typeof bkg === 'string') {
			this.bkgd = bkg
		}
	}

	fromConfig(slideDef) {
		if (slideDef.bkgd) {
			this.configureBackground(slideDef.bkgd)
		}

		// STEP 2: Add all Slide Master objects in the order they were given (Issue#53)
		if (slideDef.objects && Array.isArray(slideDef.objects) && slideDef.objects.length > 0) {
			slideDef.objects.forEach((object, idx: number) => {
				let key = Object.keys(object)[0]
				if (MASTER_OBJECTS[key] && key === 'chart') {
					this.data.push(new ChartElement(object[key].type, object[key].data, object[key].opts, this.relations))
				} else if (MASTER_OBJECTS[key] && key === 'image') {
					this.data.push(new ImageElement(object[key], this.relations))
				} else if (MASTER_OBJECTS[key] && key === 'line') {
					this.data.push(new ShapeElement(BASE_SHAPES.LINE, object[key]))
				} else if (MASTER_OBJECTS[key] && key === 'rect') {
					this.data.push(new ShapeElement(BASE_SHAPES.RECTANGLE, object[key]))
				} else if (MASTER_OBJECTS[key] && key === 'text') {
					this.data.push(new TextElement(object[key].text, object[key].options, this.relations))
				} else if (MASTER_OBJECTS[key] && key === 'placeholder') {
					const placeholder = new PlaceholderTextElement(object[key].text, object[key].options, 100 + idx, this.relations)
					if (this.placeholders.has(placeholder.name)) {
						console.warn(`Duplicate placeholders with name "${placeholder.name}"`)
						return
					}
					this.placeholders.set(placeholder.name, placeholder)
					this.data.push(placeholder)
				}
			})
		}

		// STEP 3: Add Slide Numbers
		if (slideDef.slideNumber && typeof slideDef.slideNumber === 'object') {
			this.data.push(new SlideNumberElement(slideDef.slideNumber))
		}
	}

	getPlaceholder(placeholderName?: string): PlaceholderTextElement {
		if (placeholderName && this.placeholders.has(placeholderName)) {
			return this.placeholders.get(placeholderName)
		}
		return null
	}
}

export default class SlideLayouts {
	private layoutsOrder: string[]
	private layouts: Map<string, Master>
	private presLayout

	masterSlide: Master

	constructor(presLayout) {
		this.layoutsOrder = []
		this.layouts = new Map()
		this.presLayout = presLayout

		this.new(DEF_PRES_LAYOUT_NAME)
	}

	add(layoutId, newLayout) {
		if (this.layouts.has(layoutId)) {
			throw Error('Cannot redefine a layout')
		}
		this.layoutsOrder.push(layoutId)
		this.layouts.set(layoutId, newLayout)
	}

	get(layoutId) {
		return this.layouts.get(layoutId)
	}

	provide(layoutId) {
		if (layoutId && this.layouts.has(layoutId)) return this.layouts.get(layoutId)
		return this.layouts.get(DEF_PRES_LAYOUT_NAME)
	}

	new(name: string): Master {
		const newMasterLayout = new Master(name, 1000 + this.layoutsOrder.length + 1, this.presLayout)
		this.add(name, newMasterLayout)
		return newMasterLayout
	}

	newFromConfig(name: string, config): Master {
		const newMasterLayout = this.new(name)
		if (config.margin) {
			newMasterLayout.margin = config.margin
		}
		newMasterLayout.fromConfig(config)
		return newMasterLayout
	}

	asList() {
		return this.layoutsOrder.map(l => this.get(l))
	}

	forEach(arg1, arg2?) {
		return this.asList().forEach(arg1, arg2)
	}

	map(arg1, arg2) {
		return this.asList().map(arg1, arg2)
	}

	filter(arg1, arg2) {
		return this.asList().filter(arg1, arg2)
	}
}
