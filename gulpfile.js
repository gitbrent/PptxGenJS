var gulp       = require('gulp'),
    concat     = require('gulp-concat'),
    sourcemaps = require('gulp-sourcemaps'),
    ignore     = require('gulp-ignore'),
	insert     = require('gulp-insert'),
    uglify     = require('gulp-uglify'),
	fs         = require('fs');

gulp.task('default', function(){
	var APP_VER = "", APP_BLD = "";
	fs.readFileSync("dist/pptxgen.js", "utf8").split('\n')
	.forEach((line)=>{
		if ( line.indexOf('var APP_VER') > -1 ) APP_VER = line.split('=')[1].trim().replace(/\"+|\;+/gi,'');
		if ( line.indexOf('var APP_BLD') > -1 ) APP_BLD = line.split('=')[1].trim().replace(/\"+|\;+/gi,'');
	});

	gulp.src(['libs/*', 'dist/pptxgen.js'])
        .pipe(concat('pptxgen.bundle.js'))
		.pipe(uglify())
		.pipe(insert.prepend('/* PptxGenJS '+APP_VER+'-'+APP_BLD+' */\n'))
        .pipe(sourcemaps.init())
        .pipe(ignore.exclude(["**/*.map"]))
        .pipe(sourcemaps.write('./'))
        .pipe(gulp.dest('dist/'));

    gulp.src(['dist/pptxgen.js'])
        .pipe(concat('pptxgen.min.js'))
        .pipe(uglify())
		.pipe(insert.prepend('/* PptxGenJS '+APP_VER+'-'+APP_BLD+' */\n'))
        .pipe(gulp.dest('dist/'));
});
