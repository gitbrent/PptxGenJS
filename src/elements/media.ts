import { IMG_PLAYBTN } from '../core-enums'

import { getSmartParseNumber, encodeXmlEntities } from '../gen-utils'

import ElementInterface from './element-interface'

import Hyperlink from './hyperlink'
import Position from './position'

import Relations from '../relations'

export default class MediaElement implements ElementInterface {
    videoId
    mediaId
    previewId

    mediaType
    media

    position

    constructor(options, relations: Relations) {
        this.mediaType = options.type || 'audio'
        this.media = options.path || 'preencoded.mov'

        this.position = new Position({
            x: options.x || 0,
            y: options.y || 0,

            h: options.h || 2,
            w: options.w || 2,

            flipV: options.flipV,
            flipH: options.flipH,
            rotate: options.rotate
        })

        // FIXME: 20190707
        //strType = strData ? strData.split(';')[0].split('/')[0] : strType
        let extension = 'mp3'
        extension = options.data
            ? options.data.split(';')[0].split('/')[1]
            : options.path.split('.').pop()

        // STEP 4: Add this media to this Slide Rels
        // (rId/rels count spans all slides! Count all media to get next rId)
        // NOTE: rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
        if (options.type === 'online') {
            // A: Add video
            ;[this.videoId] = relations.registerMedia({
                type: 'online',
                data: 'dummy',
                path: options.path || `preencoded${extension}`,
                extn: extension,
                target: options.link
            })
        } else {
            ;[this.videoId, this.mediaId] = relations.registerMedia({
                data: options.data || '',
                type: `${options.type}/${extension}`,
                path: options.path || `preencoded${extension}`,
                extn: extension
            })
        }

        this.previewId = relations.registerImage(
            {
                data: IMG_PLAYBTN,
                path: 'preencoded.png'
            },
            'png'
        )
    }

    render(idx, presLayout) {
        let body
        if (this.mediaType === 'online') {
            body = `
			<p:nvPicPr>
				<p:cNvPr id="${this.previewId}" name="Picture ${idx + 1}"/>
				<p:cNvPicPr/>
				<p:nvPr>
					<a:videoFile r:link="rId${this.videoId}"/>
				</p:nvPr>
			</p:nvPicPr>`
        } else {
            /*
             * IMPORTANT: <p:cNvPr id="" value is critical
             * - if not the same number as preiew image rId, PowerPoint throws error!
             */
            body = `
            <p:nvPicPr>
                <p:cNvPr id="${this.previewId}" name="Picture ${idx + 1}">
                    <a:hlinkClick r:id="" action="ppaction://media"/>
                </p:cNvPr>
			    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
			    <p:nvPr>
				    <a:videoFile r:link="rId${this.videoId}"/>
				    <p:extLst>
					    <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
                            <p14:media 
                                xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" 
                                r:embed="rId${this.mediaId}"/>
					    </p:ext>
				    </p:extLst>
			    </p:nvPr>
		    </p:nvPicPr>`
        }

        return `
        <p:pic>
            ${body}
            <p:blipFill>
                <a:blip r:embed="rId${this.previewId}"/>
                ${'' /* NOTE: Preview image is required! */}
                <a:stretch><a:fillRect/></a:stretch>
            </p:blipFill>
			<p:spPr>'
				${this.position.render(presLayout)}
				<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
			</p:spPr>
		</p:pic>`
    }
}
