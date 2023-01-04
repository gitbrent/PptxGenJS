
const PptxGenJS = require('../../dist/pptxgen.cjs')

const pptx = new PptxGenJS()

const slide = pptx.addSlide()

slide.addImage({ path: 'images/image1.png', x: 0, y: 0, appearOnClick: true })
slide.addImage({ path: 'images/image2.png', x: 1, y: 1, appearOnClick: true })
slide.addImage({ path: 'images/image3.png', x: 2, y: 2, appearOnClick: true })

pptx.writeFile({ fileName: 'out.pptx' })
