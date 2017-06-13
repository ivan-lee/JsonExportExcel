var gulp = require('gulp'),
    uglify = require('gulp-uglify'),
    rename = require('gulp-rename'),
    concat = require('gulp-concat');

gulp.task('build', function() {
  return gulp.src(['src/**/*.js','node_modules/xlsx/**/xlsx.core.min.js'])
      .pipe(concat('JsonExportExcel.js'))
      .pipe(uglify({
            preserveComments: 'license' //保留所有注释
        }))
      .pipe(rename({
          suffix: '.min'
      }))
      .pipe(gulp.dest('dist/'));
});
