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
  startApp();
});


$('.Opcja01-sidebar').click(function(){
  console.dir('Opcja01');
  $(this).parent('.sidebar').find('.active').removeClass('active green');
  $(this).addClass('active green');
  HideAll(1);
});


$('.Opcja02-sidebar').click(function(){
  console.dir('Opcja02');
  $(this).parent('.sidebar').find('.active').removeClass('active green');
  $(this).addClass('active green');
  HideAll(2);
});


$('.Opcja03-sidebar').click(function(){
  console.dir('Opcja03');
  $(this).parent('.sidebar').find('.active').removeClass('active green');
  $(this).addClass('active green');
  HideAll(3);
});


$('.Opcja04-sidebar').click(function(){
  console.dir('Opcja04');
  $(this).parent('.sidebar').find('.active').removeClass('active green');
  $(this).addClass('active green');
  HideAll(4);
});



//----------------------------------------------------------------------------------------------------------------------
function startApp(){
    $('.starter').fadeOut('fast',function(){
      $('.ui.dashboard').fadeIn('fast');
    });


    HideAll(1);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function HideAll(optionNumber){
    $('.ui.Opcja01').fadeOut('fast');
    $('.ui.Opcja02').fadeOut('fast');
    $('.ui.Opcja03').fadeOut('fast');
    $('.ui.Opcja04').fadeOut('fast');

    switch(optionNumber) {
      case 1:
        $('.ui.Opcja01').fadeIn('fast');
      break;
      case 2:
        $('.ui.Opcja02').fadeIn('fast');
      break;
      case 3:
        $('.ui.Opcja03').fadeIn('fast');
      break;
      case 4:
        $('.ui.Opcja04').fadeIn('fast');
      break;
      default:
        $('.ui.Opcja01').fadeIn('fast');
      }
}
//----------------------------------------------------------------------------------------------------------------------
