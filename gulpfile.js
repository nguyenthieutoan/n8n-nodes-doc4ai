const gulp = require('gulp');

function copyIcons() {
	return gulp.src([
		'nodes/**/*.{png,svg,webp}',
		'credentials/**/*.{png,svg,webp}'
	], { base: '.' })
	.pipe(gulp.dest('dist'));
}

exports['build:icons'] = copyIcons;
