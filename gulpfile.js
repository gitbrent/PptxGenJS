var gulp       = require('gulp'),
    concat     = require('gulp-concat'),
    sourcemaps = require('gulp-sourcemaps'),
    ignore     = require('gulp-ignore'),
    uglify     = require('gulp-uglify');

gulp.task('default', function(){
	gulp.src(['libs/*', 'dist/pptxgen.js'])
        .pipe(concat('pptxgen.bundle.js'))
        .pipe(sourcemaps.init())
        .pipe(ignore.exclude(["**/*.map"]))
        .pipe(uglify())
        .pipe(sourcemaps.write('./'))
        .pipe(gulp.dest('dist/'));

    gulp.src(['dist/pptxgen.js'])
        .pipe(concat('pptxgen.min.js'))
        .pipe(uglify())
        .pipe(gulp.dest('dist/'));
});
