const pptxgen = require("../dist/pptxgen.cjs");
const fs = require("fs-extra")
const Path = require('path');

const VIDEO_FILES_PATH = 'D:\\temp\\TANK.BD720P.mp4_Slice_20211104_144423'

async function genPPT() {
  // console.log('Init...')
  const Pptxgen = pptxgen;
  const pptx = new Pptxgen();

  const list = await fs.readdir(VIDEO_FILES_PATH)

  pptx.layout = "LAYOUT_WIDE";

  for (const k in list) {
    // console.log('Add slide', k)

    const videoFileName = list[k]
    const slide = pptx.addSlide();
    const videoPath = Path.join(VIDEO_FILES_PATH, videoFileName)

    slide.addMedia({
      x: 1.0,
      y: 2.25,
      w: 5,
      h: 2.5,
      type: 'video',
      path: videoPath,
      isFsPath: true,
      // ext: 'mp4',
      // cover: imgData.base64
    })
    slide.addText([{ text: videoPath, options: { fontFace: '微软雅黑' } }], {
      x: 7.0,
      y: 1.2,
      w: '40%',
      h: 1,
      margin: 0,
      color: '000000'
    })
  }


  console.log('Saving ppt 1...')

  // await pptx.writeFile({ fileName: "output.pptx" })

  const savePath = Path.join(__dirname, 'output.pptx')

  const fileData = await pptx.stream({
    compression: false,
  })
  console.log('createWriteStream...')
  const out = fs.createWriteStream(savePath)
  fileData.pipe(out)
    .on('finish', function () {
      console.log('Success!')
    });

  // out.write(fileData)
  // out.close()
  // out.on('error', (e) => {
  //   console.error('>>> Out Error!', e)
  // })


}

genPPT()
