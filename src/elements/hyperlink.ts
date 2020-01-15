import { encodeXmlEntities } from '../gen-utils'
import Relations from '../relations'

export interface HyperLinkOptions {
    url?: string
    slide?: number
    tooltip?: string
}

export default class HyperLink {
    url?: string
    slide?: number
    tooltip?: string
    rId?: number

    constructor(
        { url, slide, tooltip }: HyperLinkOptions,
        relations: Relations
    ) {
        if (!url && !slide)
            throw "ERROR: 'hyperlink requires either `url` or `slide`'"

        this.url = url
        this.slide = slide
        this.tooltip = tooltip

        this.rId = relations.registerLink(
            this.slide ? 'slide' : 'dummy',
            encodeXmlEntities(this.url) || this.slide.toString()
        )
    }

    render() {
        if (this.url) {
            // TODO: (20170410): FUTURE-FEATURE: color (link is always blue in Keynote and PPT online, so usual text run above isnt honored for links..?)
            //runProps += '<a:uFill>'+ genXmlColorSelection('0000FF') +'</a:uFill>'; // Breaks PPT2010! (Issue#74)
            return `<a:hlinkClick r:id="rI${this.rId}" 
            invalidUrl="" 
            action="" 
            tgtFrame="" 
            tooltip="${this.tooltip ? encodeXmlEntities(this.tooltip) : ''}" 
            history="1" 
            highlightClick="0" endSnd="0"/>`
        } else if (this.slide) {
            return `<a:hlinkClick r:id="rId${
                this.rId
            }" action="ppaction://hlinksldjump" tooltip="${
                this.tooltip ? encodeXmlEntities(this.tooltip) : ''
            }"/>`
        }
    }
}
