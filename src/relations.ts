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
}
