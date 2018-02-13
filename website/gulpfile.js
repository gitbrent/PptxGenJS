/*
 * DESC: Combine/Minify CSS and take existing min JavaScript - replace head tags with content
 * WHY.: Google PageSpeed Insights scores of 99(mobile)/94(desktop) thats why!!
 */
var fs       = require('fs'),
    gulp     = require('gulp'),
    concat   = require('gulp-concat'),
    replace  = require('gulp-string-replace'),
	cleanCSS = require('gulp-clean-css');

var cssSrch1 = '<link rel="stylesheet" href="/PptxGenJS/css/main.css"/>';
var cssSrch2 = '<link rel="stylesheet" href="//cdnjs.cloudflare.com/ajax/libs/highlight.js/9.12.0/styles/hybrid.min.css"/>';
var jvsSrch1 = /\<script type="text\/javascript" src="https:\/\/cdnjs.cloudflare.com\/ajax\/libs\/highlight.*.min.js"\>\<\/script\>/;
var jvsSrch2 = 'pptxgen.bundle.js">';

/* ========== */
var arrDeployTasks = ['deploy-html','deploy-index','deploy-img','deploy-help','deploy-sitemap'];

gulp.task('deploy-html', ()=>{
	return gulp.src('./build/PptxGenJS/docs/*.html').pipe(gulp.dest('../docs/'));
});

gulp.task('deploy-index', ()=>{
	return gulp.src('../index.perf.html', {base:'./'}).pipe(gulp.dest('../index.html'));
});

gulp.task('deploy-img', ()=>{
	return gulp.src('./build/PptxGenJS/img/*.*').pipe(gulp.dest('../img/'));
});

gulp.task('deploy-help', ()=>{
	return gulp.src('./build/PptxGenJS/help.html').pipe(gulp.dest('../'));
});

gulp.task('deploy-sitemap', ()=>{
	return gulp.src('./build/PptxGenJS/sitemap.xml').pipe(gulp.dest('../'));
});

gulp.task('deploy', arrDeployTasks, ()=>{
	console.log('Deploy tasks run:'+ arrDeployTasks.length );
	console.log('DONE!');
});

/* ========== */

gulp.task('min-css', ()=>{
	// STEP 1: Inline both css files
	return gulp.src(['../css/hybrid.min.css', './build/PptxGenJS/css/main.css'])
		.pipe(concat('style.bundle.css'))
		.pipe(cleanCSS())
		.pipe(gulp.dest('../css'))
		.pipe(gulp.src('../css/style.bundle.css'));
});

gulp.task('default', ['min-css'], ()=>{
	// STEP 2: Grab newly combined styles
	var strMinCss = fs.readFileSync('../css/style.bundle.css', 'utf8');
	var strMinJvs = fs.readFileSync('../js/highlight.min.js', 'utf8');
	//console.log('>> `style.bundle.css` lines = '+ strMinCss.split('\n').length);
	//console.log('>> `highlight.min.js` lines = '+ strMinJvs.split('\n').length);

	// STEP 3: Replace head tags with inline style/javascript
	gulp.src('build/PptxGenJS/index.html')
	.pipe(replace(cssSrch1, '', {logs:{ enabled:true }}))
	.pipe(replace(cssSrch2, '\n<style>'+ strMinCss +'</style>\n', {logs:{ enabled:false }}))
	.pipe(replace(jvsSrch1, '<script>'+ strMinJvs +'</script>\n', {logs:{ enabled:false }}))
	.pipe(replace(jvsSrch2, 'pptxgen.bundle.js" async>', {logs:{ enabled:true }}))
	.pipe(concat('index.perf.html'))
	.pipe(gulp.dest('../'));
});
