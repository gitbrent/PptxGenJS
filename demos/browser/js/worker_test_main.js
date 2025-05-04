// demos/browser/js/worker_test_main.js
document.addEventListener('DOMContentLoaded', () => {
	const generateButton = document.getElementById('generatePptWorker');
	const statusDiv = document.getElementById('workerStatus');

	if (!generateButton || !statusDiv) {
		console.error('Required HTML elements not found!');
		statusDiv.textContent = 'Error: Could not find necessary HTML elements.';
		return;
	}

	// Create the Web Worker instance
	// The path is relative to the HTML file location
	const pptxWorker = new Worker('./js/pptxgenjs_worker.js');
	// const pptxWorker = new Worker('./js/test_worker.js'); // TESTING ONLY

	// Listen for messages *from* the worker
	pptxWorker.onmessage = function(event) {
		console.log('Main thread received message from worker:', event.data);
		const message = event.data;

		if (message.type === 'status') {
			statusDiv.textContent = `Status: ${message.message}`;
			// Disable button while working
			generateButton.disabled = true;
		} else if (message.type === 'blobGenerated') {
			const pptBlob = message.data;
			//statusDiv.textContent = 'Status: Presentation generated successfully!';

			// Use FileSaver.js to save the blob
			// You might need to include FileSaver.js in worker_test.html
			if (typeof saveAs === 'function') {
				saveAs(pptBlob, 'worker_demo.pptx');
				statusDiv.textContent += ' Downloading blob...';
			} else {
				statusDiv.textContent += ' Generated, but FileSaver.js not available to download.';
				console.error('FileSaver.js not found. Cannot save the generated blob.');
			}

			generateButton.disabled = false;
		} else if (message.type === 'buffGenerated') {
			const pptBlob = new Blob(
				[message.buffer],
				{ type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }
			)
			saveAs(pptBlob, 'worker_demo.pptx')
			statusDiv.textContent += ' Downloading arrayBuffer...';
		} else if (message.type === 'error') {
			statusDiv.textContent = `Error: ${message.message}`;
			generateButton.disabled = false; // Re-enable button
			console.error('Error received from worker:', message.message);
		}
	};

	// Handle potential errors from the worker itself (e.g. script loading failed)
	pptxWorker.onerror = function(error) {
		statusDiv.textContent = `Worker Error: ${error.message}`;
		generateButton.disabled = false;
		console.error('Web Worker encountered an error:', error);
	};

	// Add event listener to the button
	generateButton.addEventListener('click', () => {
		statusDiv.textContent = 'Status: Sending request to worker...';
		generateButton.disabled = true; // Disable button immediately

		// Send a message to the worker to start generation
		pptxWorker.postMessage({ type: 'generatePpt' });
	});

	statusDiv.textContent = 'Status: Page loaded, worker initialized.';
});
