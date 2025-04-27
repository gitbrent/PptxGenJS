// demos/browser/js/pptxgenjs_worker.js
console.log('Simple worker script loaded successfully!');

self.onmessage = function(event) {
	console.log('Simple worker received message:', event.data);
	self.postMessage({ type: 'testSuccess', message: 'Worker received and responded!' });
};

self.postMessage({ type: 'status', message: 'Simple worker is alive.' });
