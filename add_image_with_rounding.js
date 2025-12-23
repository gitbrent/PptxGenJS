/*
 * NAME: add_image_with_rounding.js
 * DESC: Adds an image to a slide with rounding enabled
 * USAGE: node add_image_with_rounding.js
 */

import pptxgen from "pptxgenjs";

const pptx = new pptxgen();

// Create a new slide
const slide = pptx.addSlide();


// Add the image with rounding enabled
slide.addImage({
	path: "https://files.chroniclehq.com/card-background-v2/thumbnail/v2-gradient-01-light.jpg",
	x: 0.528,
	y: 1.051,
	w: 2.479,
	h: 1.76,
	rounding: true,
	rectRadius: 0.2  // Optional: Set corner radius (0.0 = no rounding, 1.0 = maximum rounding)
});

slide.addImage({
	path: "https://files.chroniclehq.com/card-background-v2/thumbnail/v2-gradient-13-light.jpg",
	x: 3.87,
	y: 1.051,
	w: 2.479,
	h: 1.76,
	rounding: true,
	rectRadius: 0.2  // Optional: Set corner radius (0.0 = no rounding, 1.0 = maximum rounding)
});

slide.addImage({
	path: "atom.svg", // Replace with your SVG file path
	x: 0.628, // Position after "card" text (adjust as needed)
	y: 1.295,
	w: 0.15, // Icon width (adjust size as needed)
	h: 0.15  // Icon height (adjust size as needed)
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
	rectRadius: 0.2  // Optional: Set corner radius (0.0 = no rounding, 1.0 = maximum rounding)
});

slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.528,        // Horizontal position in inches
    y: 3.212,        // Vertical position in inches
    w: 2.5,        // Width in inches
    h: 2.0,        // Height in inches
    fill: { color: "000000", transparency: 92 }, // rgba(0,0,0,0.04) = black with 96% transparency
    rectRadius: 0.2  // Corner radius: 0 = sharp corners, 1 = fully rounded
  });

  slide.addImage({
	path: "acorn.svg", // Replace with your SVG file path
	x: 0.628, // Position after "card" text (adjust as needed)
	y: 3.3,
	w: 0.15, // Icon width (adjust size as needed)
	h: 0.15  // Icon height (adjust size as needed)
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




// Save the presentation
const exportName = "Image_With_Rounding";
pptx.writeFile({ fileName: exportName })
	.then((fileName) => {
		console.log(`Presentation exported: ${fileName}`);
	})
	.catch((err) => {
		console.error(`ERROR: ${err}`);
	});

