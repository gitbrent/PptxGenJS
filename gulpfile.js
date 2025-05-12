/*\
|*| DESC: Combine/Minify CSS and take existing min JavaScript - replace head tags with content
|*| WHY.: Google PageSpeed Insights [Lighthouse] scores of 100(mobile)/100(desktop) thats why!!
\*/
"use strict";
let gulp = require("gulp");

// DOCS
gulp.task("deploy-assets", () => gulp.src("./build/assets/*/*").pipe(gulp.dest("./assets")));
gulp.task("deploy-docs", () => gulp.src("./build/docs/**").pipe(gulp.dest("./docs/")));

// PAGES
gulp.task("deploy-html", () => gulp.src("./build/*.html").pipe(gulp.dest("./")));
gulp.task("deploy-demo", () => gulp.src("./build/demo**/**", {encoding: false}).pipe(gulp.dest("./")));
gulp.task("deploy-html2pptx", () => gulp.src("./build/html2pptx**/**").pipe(gulp.dest("./")));
gulp.task("deploy-license", () => gulp.src("./build/license**/**").pipe(gulp.dest("./")));
gulp.task("deploy-privacy", () => gulp.src("./build/privacy**/**").pipe(gulp.dest("./")));
gulp.task("deploy-sponsor", () => gulp.src("./build/sponsor**/**").pipe(gulp.dest("./")));

// IMG & SITEMAP
gulp.task("deploy-img", () => gulp.src("./build/img/*.*", {encoding: false}).pipe(gulp.dest("./img/")));
gulp.task("deploy-sitemap", () => gulp.src("./build/sitemap.xml").pipe(gulp.dest("./")));

// Build/Deploy
gulp.task(
	"default",
	gulp.parallel(
		"deploy-assets",
		"deploy-docs",
		"deploy-html",
		"deploy-demo",
		"deploy-html2pptx",
		"deploy-license",
		"deploy-privacy",
		"deploy-sponsor",
		"deploy-img",
		"deploy-sitemap"
	),
	() => {
		console.log("...gulp build done!");
	}
);
