import {
    IImageOpts,
    IMediaOpts,
    IShape,
    IShapeOptions,
    ITableOptions,
    IText,
    ITextOpts,
    TableRow
} from '../core-interfaces'

import Relations from '../relations'

import ElementInterface from './element-interface'
import TextElement from './text'
import ShapeElement from './simple-shape'
import ImageElement from './image'
import ChartElement from './chart'
import SlideNumberElement from './slide-number'
import MediaElement from './media'

import Position from './position'

type PositionnedElement = ElementInterface & { position: Position }

export default class GroupElement implements ElementInterface {
    data: PositionnedElement[] = []
    relations: Relations

    constructor(relations: Relations) {
        this.relations = relations
    }

    addSlideNumber(value): GroupElement {
        this.data.push(new SlideNumberElement(value, this.relations))
        return this
    }

    addChart(type, data, options): GroupElement {
        this.data.push(new ChartElement(type, data, options, this.relations))
        return this
    }

    addImage(options: IImageOpts): GroupElement {
        if (!options.path && !options.data) {
            console.error(
                "ERROR: `addImage()` requires either 'data' or 'path' parameter!"
            )
            return null
        } else if (
            options.data &&
            options.data.toLowerCase().indexOf('base64,') === -1
        ) {
            console.error(
                "ERROR: Image `data` value lacks a base64 header! Ex: 'image/png;base64,NMP[...]')"
            )
            return null
        }

        this.data.push(new ImageElement(options, this.relations))
        return this
    }

    addMedia(options: IMediaOpts): GroupElement {
        this.data.push(new MediaElement(options, this.relations))
        return this
    }

    addShape(shape: IShape, options?: IShapeOptions): GroupElement {
        this.data.push(new ShapeElement(shape, options))
        return this
    }

    addText(text: string | IText[], options?: ITextOpts): GroupElement {
        this.data.push(new TextElement(text, options, this.relations))
        return this
    }

    render(idx, presLayout) {
        const xPos = this.data
            .map(x => x.position && x.position.xPos(presLayout))
            .filter(x => !!x)
        const minX = Math.min(...xPos.map(([x0]) => x0))
        const maxX = Math.max(...xPos.map(([, x1]) => x1))

        const yPos = this.data
            .map(y => y.position && y.position.yPos(presLayout))
            .filter(y => !!y)
        const minY = Math.min(...yPos.map(([y0]) => y0))
        const maxY = Math.max(...yPos.map(([, y1]) => y1))

        return `
      <p:grpSp>
        <p:nvGrpSpPr>
          <p:cNvPr id="${idx + 1}" name="Group ${idx}">
            <a:extLst>
              <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
                <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{34B72804-196E-FE49-B905-024A3D6F4AA4}"/>
              </a:ext>
            </a:extLst>
          </p:cNvPr>
          <p:cNvGrpSpPr/>
          <p:nvPr/>
        </p:nvGrpSpPr>
        <p:grpSpPr>
          <a:xfrm>
            <a:off x="${minX}" y="${minY}"/>
            <a:ext cx="${maxX - minX}" cy="${maxY - minY}"/>
            <a:chOff x="${minX}" y="${minY}"/>
            <a:chExt cx="${maxX - minX}" cy="${maxY - minY}"/>
          </a:xfrm>
        </p:grpSpPr>
        ${this.data
            // TODO: find a better indexing method
            .map((d, i) => d.render(idx * 1000 + i, presLayout))
            .join('')}
      </p:grpSp>`
    }
}
