function loadSection(id, file) {
	fetch(file)
		.then((response) => {
			if (!response.ok) {
				throw new Error(`Failed to load ${file}: ${response.statusText}`);
			}
			return response.text();
		})
		.then((html) => {
			document.getElementById(id).outerHTML = html;
			//console.log(`Loaded ${file} into ${id}`);
		})
		.catch((error) => console.error(error));
}

// Load sections
loadSection('navbar', './html/navbar.html');
loadSection('header', './html/header.html');
loadSection('navtabs', './html/navtabs.html');
loadSection('tab-intro', './html/tab-intro.html');
loadSection('tab-html2pptx', './html/tab-html2pptx.html');
loadSection('tab-charts', './html/tab-charts.html');
loadSection('tab-images', './html/tab-images.html');
loadSection('tab-shapes', './html/tab-shapes.html');
loadSection('tab-tables', './html/tab-tables.html');
loadSection('tab-masters', './html/tab-masters.html');
