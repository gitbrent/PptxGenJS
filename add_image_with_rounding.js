
import pptxgen from "pptxgenjs";

const pptx = new pptxgen();

const slide = pptx.addSlide();


slide.addImage({
	path: "https://files.chroniclehq.com/card-background-v2/thumbnail/v2-gradient-01-light.jpg",
	x: 0.528,
	y: 1.051,
	w: 2.479,
	h: 1.76,
	rounding: true,
	rectRadius: 0.2 
});

slide.addImage({
	path: "https://files.chroniclehq.com/card-background-v2/thumbnail/v2-gradient-13-light.jpg",
	x: 3.87,
	y: 1.051,
	w: 2.479,
	h: 1.76,
	rounding: true,
	rectRadius: 0.2
});

slide.addImage({
	path: "atom.svg", 
	x: 0.628, 
	y: 1.295,
	w: 0.15, 
	h: 0.15  
});

slide.addText("Card 1", {
	x: 0.528,
	y: 1.555,
	w: 0.751,
	h: 0.208,
	fontSize: 14,
	color: "050505",
    align: "left"
});

slide.addText("Lorem ipsum dolor sit amet", {
	x: 0.528,
	y: 1.759,
	w: 2.114,
	h: 0.208,
	fontSize: 12,
	color: "050505",
    align: "left"
});

slide.addText("Ut enim ad minim", {
	x: 0.528,
	y: 2.295,
	w: 1.48,
	h: 0.188,
	fontSize: 12,
	color: "050505",
    align: "left"
});

slide.addText("Lorem ipsum dolor sit amet, consectetur adipiscing elit.", {
	x: 4.051,
	y: 1.862,
	w: 2.114,
	h: 0.188,
	fontSize: 12,
	color: "050505",
    align: "left"
});

slide.addImage({
	path: "https://files.chroniclehq.com/card-background-v2/thumbnail/v2-image-architecture-01.jpg",
	x: 6.893,
	y: 1.169,
	w: 1.992,
	h: 2.755,
    sizing: {
        type: "crop",
        w: 1.992,
        h: 2.755
    },
	rounding: true,
	rectRadius: 0.2 
});

slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.528,        // Horizontal position in inches
    y: 3.212,        // Vertical position in inches
    w: 2.5,        // Width in inches
    h: 2.0,        // Height in inches
    fill: { color: "000000", transparency: 92 },
    rectRadius: 0.2 
  });

  slide.addImage({
	path: "acorn.svg", 
	x: 0.628, 
	y: 3.3,
	w: 0.15, 
	h: 0.15  
});

slide.addText("Card 3", {
	x: 0.528,
	y: 3.669,
	w: 0.751,
	h: 0.208,
	fontSize: 14,
	color: "050505",
    align: "left"
});

slide.addText("Lorem ipsum dolor sit amet", {
	x: 0.528,
	y: 3.925,
	w: 2.114,
	h: 0.208,
	fontSize: 12,
	color: "050505",
    align: "left"
});

slide.addText("Ut enim ad minim", {
	x: 0.528,
	y: 4.751,
	w: 1.48,
	h: 0.188,
	fontSize: 12,
	color: "050505",
    align: "left"
});




const exportName = "Image_With_Rounding";
pptx.writeFile({ fileName: exportName })
	.then((fileName) => {
		console.log(`Presentation exported: ${fileName}`);
	})
	.catch((err) => {
		console.error(`ERROR: ${err}`);
	});

