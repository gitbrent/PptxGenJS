import { SLIDE_OBJECT_TYPES } from './core-enums'
import XML_HEADER from './templates/xml-header'

const relationship = (id, rType, target, other: [string, string][] = []) => {
    return [
        `<Relationship Id="rId${id}"`,
        ` Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${rType}"`,
        ` Target="${target}"`,
        other.map((k, v) => ` ${k}="${v}"`).join(''),
        '/>'
    ].join('')
}

const relationship2007 = (id, rType, target) => {
    return [
        `<Relationship Id="rId${id}"`,
        ` Type="http://schemas.microsoft.com/office/2007/relationships/${rType}"`,
        ` Target="${target}"`,
        '/>'
    ].join('')
}

export default class Relations {
    rels = []
    relsChart = []
    relsMedia = []

    registerLink(data, target) {
        const relId =
            this.rels.length + this.relsChart.length + this.relsMedia.length + 1
        this.rels.push({
            type: SLIDE_OBJECT_TYPES.hyperlink,
            data,
            rId: relId,
            Target: target
        })

        return relId
    }

    registerImage({ path, data = '' }, extension, fromSvgSize = false) {
        // (rId/rels count spans all slides! Count all images to get next rId)
        const relId =
            this.rels.length + this.relsChart.length + this.relsMedia.length + 1

        const Target = `../media/image-${Math.random()}.${extension}`
        const mediaConfig = {
            rId: relId,
            type: `image/${extension === 'svg' ? 'svg+xml' : extension}`,

            path: path,
            data: data,

            extn: extension,
            Target,

            isSvgPng: false,
            svgSize: null
        }

        if (fromSvgSize) {
            mediaConfig.isSvgPng = true
            mediaConfig.svgSize = fromSvgSize
        }

        this.relsMedia.push(mediaConfig)
        return relId
    }

    registerChart(globalId, options, data) {
        const chartRid = this.relsChart.length + 1

        this.relsChart.push({
            rId: chartRid,
            data,
            opts: options,
            type: options.type,
            globalId: globalId,
            fileName: 'chart' + globalId + '.xml',
            Target: '/ppt/charts/chart' + globalId + '.xml'
        })

        return chartRid
    }

    registerMedia({ path, type, extn, data = null, target = null }) {
        const relId =
            this.rels.length + this.relsChart.length + this.relsMedia.length + 1

        let config = {
            rId: relId,
            type,
            path,
            extn,
            data,
            Target: null
        }
        if (type === 'online') {
            config.Target = target
            this.relsMedia.push(config)
            return [relId]
        }

        const Target = `../media/image-${Math.random()}.${extn}`
        config.Target = Target
        this.relsMedia.push(config)

        const relId2 = relId + 1
        const config2 = { ...config, Target, rId: relId2 }
        this.relsMedia.push(config2)

        return [relId, relId2]
    }

    render(defaultRels) {
        let lastRid = 0 // stores maximum rId used for dynamic relations
        let strXml = [
            XML_HEADER,
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        ].join('')

        this.rels.forEach(rel => {
            lastRid = Math.max(lastRid, rel.rId)
            if (rel.type.toLowerCase().indexOf('hyperlink') > -1) {
                if (rel.data === 'slide') {
                    strXml += relationship(
                        rel.rId,
                        'slide',
                        `slide${rel.Target}.xml`
                    )
                } else {
                    strXml += relationship(rel.rId, 'hyperlink', rel.Target, [
                        ['TargetMode', 'External']
                    ])
                }
            } else if (rel.type.toLowerCase().indexOf('notesslide') > -1) {
                strXml += relationship(rel.rId, 'notesSlide', rel.Target)
            }
        })

        this.relsChart.forEach(rel => {
            lastRid = Math.max(lastRid, rel.rId)
            strXml += relationship(rel.rId, 'chart', rel.Target)
        })

        this.relsMedia.forEach(rel => {
            lastRid = Math.max(lastRid, rel.rId)
            if (rel.type.toLowerCase().indexOf('image') > -1) {
                strXml += relationship(rel.rId, 'image', rel.Target)
            } else if (rel.type.toLowerCase().indexOf('audio') > -1) {
                // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
                if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                    strXml += relationship2007(rel.rId, 'media', rel.Target)
                else strXml += relationship(rel.rId, 'audio', rel.Target)
            } else if (rel.type.toLowerCase().indexOf('video') > -1) {
                // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
                if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                    strXml += relationship2007(rel.rId, 'media', rel.Target)
                else strXml += relationship(rel.rId, 'video', rel.Target)
            } else if (rel.type.toLowerCase().indexOf('online') > -1) {
                // As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
                if (strXml.indexOf(' Target="' + rel.Target + '"') > -1)
                    strXml += relationship2007(rel.rId, 'image', rel.Target)
                else
                    strXml += relationship(rel.rId, 'video', rel.Target, [
                        ['TargetMode', 'External']
                    ])
            }
        })

        defaultRels.forEach((rel, idx) => {
            strXml += relationship(lastRid + idx + 1, rel.type, rel.target)
        })

        strXml += '</Relationships>'
        return strXml
    }
}
