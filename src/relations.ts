import { SLIDE_OBJECT_TYPES } from './core-enums'

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

    registerImage({ path, data = '' }, extension, fromSvgSize) {
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
}
