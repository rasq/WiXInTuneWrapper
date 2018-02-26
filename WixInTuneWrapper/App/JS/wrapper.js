const remote      = require('electron').remote;
const dialog      = remote.dialog;
const fs          = require('fs');
const session     = remote.session;
const needle      = require('needle');
const prompt      = require('dialogs')(opts={});
const mkdirp      = require('mkdirp');
const homedir     = require('os').homedir();
const sanitize    = require("sanitize-filename");
var Downloader    = require('mt-files-downloader');
var shell         = require('electron').shell;

$('.ui.dropdown').dropdown();


jQuery.expr[':'].contains = function(a, i, m) {
  return jQuery(a).text().toUpperCase().indexOf(m[3].toUpperCase()) >= 0;
};


$('.ui.loading .form').submit((e)=>{
    console.dir('bvbvbvbvvbvb');
});

$('.start-button').click(function(){
  $('.content .ui.wrappers,.content .ui.dimmer').hide();
  $('.content .ui.dashboard').show();
  $(this).parent('.sidebar').find('.active').removeClass('active red');
  $(this).addClass('active red');
  //loadSettings();
    console.dir('dadadadadadada');
});
