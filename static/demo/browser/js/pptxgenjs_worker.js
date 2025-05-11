// demos/browser/js/pptxgenjs_worker.js

// IMPORTANT: You need to load pptxgenjs within the worker.
// Assuming your built pptxgen.js is available relative to the worker script.
// Adjust the path as needed based on your project structure.
try {
	importScripts('./pptxgen.bundle.js');
	console.log('pptxgenjs loaded successfully in worker.');
} catch (e) {
	console.error('Failed to load pptxgenjs in worker:', e);
	// You might want to post an error message back to the main thread
}

// Listen for messages from the main thread
self.onmessage = async function(event) {
	console.log('Worker received message:', event.data);

	const message = event.data;

	if (message.type === 'generatePpt') {
		try {
			// Inform the main thread that generation is starting
			self.postMessage({ type: 'status', message: 'Generating presentation...' });

			// *** pptxgenjs code runs here ***
			let pptx = new PptxGenJS();
			let slide = pptx.addSlide();
			slide.addText(
				'👷 Hello from Web Worker!',
				{ x: 1, y: 1, w: 8, h: 1, fontSize: 24, fill: { color: 'FFFF00' } }
			);
			slide.addText(
				`Generated at: ${new Date().toLocaleString()}`,
				{ x: 1, y: 2, w: 4, h: 0.5, fontSize: 14 }
			);
			slide.addText(
				`Library version: ${pptx.version}`,
				{ x: 5, y: 2, w: 4, h: 0.5, fontSize: 14 }
			);
			// Test with an image from a URL as Issue #1354 called this out)
			slide.addImage({
				path: "https://raw.githubusercontent.com/gitbrent/PptxGenJS/master/demos/common/images/krita_square.jpg",
				x: 1.0, y: 3.1, w: 2.0, h: 2.0
			});
			slide.addText("<-- test image via `path` URL", { x: 3.1, y: 4.7, w: 4, h: 0.5, fontSize: 14, color: '0000FF' });

			// Generate the presentation as a Blob or ArrayBuffer to send back
			// Blob is often easiest for saving/downloading in the main thread
			const blob = await pptx.write('blob');
			self.postMessage({ type: 'blobGenerated', data: blob })
			// COMMENTED OUT: This is a test to see if the arrayBuffer works
			/*
			const buffer = await pptx.write('arraybuffer');
			self.postMessage({ type: 'buffGenerated', buffer }, [buffer])
			*/
		} catch (error) {
			console.error('Error generating presentation in worker:', error);
			// Send an error message back to the main thread
			self.postMessage({ type: 'error', message: error.message || 'An error occurred during generation.' });
		}
	}
};

console.log('Web Worker script loaded.');
