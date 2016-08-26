var gulp = require('gulp');
var watch= require('gulp-watch');
var uglify= require('gulp-uglify');

var exec = require("child_process").exec;
var config={uglify:false}

 
gulp.task('watch', function (cb) {
    	exec("mocha", function (error, stdout, stderr) {
		if (stdout) {
			console.log("stdout: " + stdout);
		}
		if (stderr) {
			console.log("stderr: " + stderr);
		}
		if (error !== null) {
			console.log("exec error: " + error);
		}
	});
});



 
 
gulp.task('default');
