var gulp       = require('gulp'),
	concat     = require('gulp-concat'),
	sourcemaps = require('gulp-sourcemaps'),
	ignore     = require('gulp-ignore'),
	insert     = require('gulp-insert'),
	uglify     = require('gulp-uglify'),
	fs         = require('fs');

var APP_VER = "", APP_BLD = "";
gulp.task('get-vars', ()=>{
	fs.readFileSync("dist/pptxgen.js", "utf8").split('\n')
	.forEach((line)=>{
		if ( line.indexOf('var APP_VER') > -1 ) APP_VER = line.split('=')[1].trim().replace(/\"+|\;+/gi,'');
		if ( line.indexOf('var APP_BLD') > -1 ) APP_BLD = line.split('=')[1].trim().replace(/\"+|\;+/gi,'');
	});
	return gulp.src('README.md');
});
gulp.task('deploy-bundle', ()=>{
	return gulp.src(['libs/*', 'dist/pptxgen.js'])
		.pipe(concat('pptxgen.bundle.js'))
		.pipe(uglify())
		.pipe(insert.prepend('/* PptxGenJS '+APP_VER+'-'+APP_BLD+' */\n'))
		.pipe(sourcemaps.init())
		.pipe(ignore.exclude(["**/*.map"]))
		.pipe(sourcemaps.write('./'))
		.pipe(gulp.dest('dist/'));
});
gulp.task('deploy-min', ()=>{
	return gulp.src(['dist/pptxgen.js'])
		.pipe(concat('pptxgen.min.js'))
		.pipe(uglify())
		.pipe(insert.prepend('/* PptxGenJS '+APP_VER+'-'+APP_BLD+' */\n'))
		.pipe(gulp.dest('dist/'));
});

// Build/Deploy
gulp.task('default', gulp.series('get-vars', gulp.parallel('deploy-bundle','deploy-min')), ()=>{
	console.log('Done');
});
