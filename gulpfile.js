const pkg = require("./package.json");
const rollup = require("rollup");
const rollupTypescript = require("rollup-plugin-typescript2");
const { watch, series } = require("gulp");
const gulp = require("gulp"),
  concat = require("gulp-concat"),
  sourcemaps = require("gulp-sourcemaps"),
  ignore = require("gulp-ignore"),
  insert = require("gulp-insert"),
  deleteLines = require("gulp-delete-lines"),
  uglify = require("gulp-uglify");

gulp.task("build", () => {
  return rollup
    .rollup({
      input: "./src/pptxgen.ts",
      external: [
        ...Object.keys(pkg.dependencies || {}),
        ...Object.keys(pkg.peerDependencies || {})
      ],
      plugins: [rollupTypescript()]
    })
    .then(bundle => {
      return bundle.write({
        file: "./src/bld/pptxgen.gulp.js",
        format: "iife",
        name: "PptxGenJS",
        globals: {
            jszip: 'JSZip'
        },
        sourcemap: true,
		definitions: true
      });
    });
});

gulp.task("clean", () => {
  return gulp
    .src(["./src/bld/pptxgen.gulp.js"])
    .pipe(concat("pptxgen.min.js"))
    .pipe(uglify())
    .pipe(
      insert.prepend(
        "/* PptxGenJS " +
          pkg.version +
          " @ " +
          new Date().toISOString() +
          " */\n"
      )
    )
    .pipe(sourcemaps.init())
    .pipe(ignore.exclude(["**/*.map"]))
    .pipe(sourcemaps.write("./"))
    .pipe(gulp.dest("./dist/"));
});

// Build/Deploy
gulp.task("default", gulp.series("build", "clean"), () => {
  console.log("... dist/pptxgen.min.js done!");
});

// Watch
exports.default = function() {
  watch("src/*.ts", series("build", "clean"));
};
