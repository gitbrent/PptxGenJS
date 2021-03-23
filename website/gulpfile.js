/*\
|*| DESC: Combine/Minify CSS and take existing min JavaScript - replace head tags with content
|*| WHY.: Google PageSpeed Insights [Lighthouse] scores of 100(mobile)/100(desktop) thats why!!
\*/
"use strict";
let gulp = require("gulp");

// TASKS: Deploy
gulp.task("deploy-assets", () => {
	return gulp.src("./build/assets/*/*").pipe(gulp.dest("../assets"));
});
gulp.task("deploy-docs", () => {
	return gulp.src("./build/docs/**").pipe(gulp.dest("../docs/"));
});
gulp.task("deploy-html", () => {
	return gulp.src("./build/*.html").pipe(gulp.dest("../"));
});
gulp.task("deploy-license", () => {
	return gulp.src("./build/license**/**").pipe(gulp.dest("../"));
});
gulp.task("deploy-privacy", () => {
	return gulp.src("./build/privacy/**/**").pipe(gulp.dest("../"));
});
gulp.task("deploy-img", () => {
	return gulp.src("./build/img/*.*").pipe(gulp.dest("../img/"));
});
gulp.task("deploy-sitemap", () => {
	return gulp.src("./build/sitemap.xml").pipe(gulp.dest("../"));
});

// Build/Deploy
gulp.task("default", gulp.parallel("deploy-assets", "deploy-docs", "deploy-html", "deploy-license", "deploy-privacy", "deploy-img", "deploy-sitemap"), () => {
	console.log("Done");
});
