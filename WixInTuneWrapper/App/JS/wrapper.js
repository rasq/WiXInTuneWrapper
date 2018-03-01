const remote      = require('electron').remote;
const dialog      = remote.dialog;
const fs          = require('fs');
const session     = remote.session;
const needle      = require('needle');
const prompt      = require('dialogs')(opts={});
const mkdirp      = require('mkdirp');
const homedir     = require('os').homedir();
const sanitize    = require("sanitize-filename");
const uuidV4      = require('uuid/v4');
const uuidValidate = require('uuid-validate');
//var Downloader    = require('mt-files-downloader');
var shell         = require('electron').shell;


var customActionTPL = "<?xml version='1.0'?>" +
"<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'>" +
   '<Fragment>' +
      "<CustomAction Id='FooAction' BinaryKey='FooBinary' DllEntry='FooEntryPoint' Execute='immediate' Return='check'/>" +
      "<Binary Id='FooBinary' SourceFile='foo.dll'/>" +
   "</Fragment>" +
"</Wix>"


var installExecuteSeqTPLH = "<InstallExecuteSequence>" +
         "<Custom Action='actionName' After='InstallFiles'/>" +
      "</InstallExecuteSequence>"








$('.getFolder-button').click(function(){
  selectDirectory();
  //document.getElementById('selectFolderA').click();
});

$('.getIco-button').click(function(){
  var dataIco = document.getElementById("selectIcoA");
    dataIco.click(function(){
    });
});


function updateIcoVar(fileInput) {
        var files = fileInput.files;
        for (var i = 0; i < files.length; i++) {
            var file = files[i];
            var imageType = /image.*/;
            if (!file.type.match(imageType)) {
                continue;
            }
            var img = document.getElementById("thumbnil");
            img.file = file;
            //console.log('value ' + file.path);
            var reader = new FileReader();
            reader.onload = (function(aImg) {
                return function(e) {
                    aImg.src = e.target.result;
                };
            })(img);
            reader.readAsDataURL(file);
        }
}

//----------------------------------------------------------------------------------------------------------------------
$('.generateWXS-button').click(function(){
  harvestDataFromForms(1);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.start-button').click(function(){
  startApp('.Opcja01-sidebar');
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.ui.green.button.PC').click(function(){
  document.getElementsByName("ProductCode")[0].value = '{' + uuidV4() + '}';
});

$('.ui.green.button.UC').click(function(){
  document.getElementsByName("UpgradeCode")[0].value = '{' + uuidV4() + '}';
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.Opcja01-sidebar').click(function(){
  HideAll(this, '.ui.Opcja01');
});


$('.Opcja02-sidebar').click(function(){
  HideAll(this, '.ui.Opcja02');
});


$('.Opcja03-sidebar').click(function(){
  HideAll(this, '.ui.Opcja03');
});


$('.Opcja04-sidebar').click(function(){
  HideAll(this, '.ui.Opcja04');
});


$('.Opcja05-sidebar').click(function(){
  HideAll(this, '.ui.Opcja05');
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.ui.checkbox.left.A').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'), 'CB_CA');
});

$('.ui.checkbox.left.B').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'), 'CB_LC');
});

$('.ui.checkbox.left.C').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'), 'CB_Shortcuts');
});

$('.ui.checkbox.left.D').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'), 'CB_Reg');
});

$('.ui.checkbox.left.A').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'), 'CB_CA');
});

/*$('.ui.checkbox.left.P').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'));
});

$('.ui.checkbox.left.L').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'));
});

$('.ui.checkbox.left.M').click(function(){
  CheckBoxMSIActions(!$(this).checkbox('is checked'));
});*/
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function startApp(initTab){
    $('.starter').fadeOut('fast',function(){
      $('.ui.dashboard').fadeIn('fast');
    });

    HideAll(initTab, '.ui.Opcja01');
    HideAllCheckBoxMSI();
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function HideAll(optionName, tabName){
    $('.ui.Opcja01').fadeOut('fast');
    $('.ui.Opcja02').fadeOut('fast');
    $('.ui.Opcja03').fadeOut('fast');
    $('.ui.Opcja04').fadeOut('fast');

    $(tabName).fadeIn('fast');

    $(optionName).parent('.sidebar').find('.active').removeClass('active green');
    $(optionName).addClass('active green');
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function HideAllCheckBoxMSI(){
    $('.ui.checkbox.left.P').checkbox('set unchecked');
    $('.ui.checkbox.left.L').checkbox('set unchecked');
    $('.ui.checkbox.left.M').checkbox('set unchecked');

    $('.ui.checkbox.left.A').checkbox('set unchecked');
    $('.ui.checkbox.left.B').checkbox('set unchecked');
    $('.ui.checkbox.left.C').checkbox('set unchecked');
    $('.ui.checkbox.left.D').checkbox('set unchecked');

    $('.CB_CA').fadeOut('fast');
    $('.CB_LC').fadeOut('fast');
    $('.CB_Shortcuts').fadeOut('fast');
    $('.CB_Reg').fadeOut('fast');
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function CheckBoxMSIActions(state, action){
  if (state == true) {
    $('.' + action).fadeIn('fast');
  } else {
    $('.' + action).fadeOut('fast');
  }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function harvestDataFromForms(pageID){
  if (pageID == 1) {
    formConfigMSI();
  } else if (pageID == 2) {

  } else if (pageID == 3) {

  }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function validateAppVersion(number){
  var tmpVar = number.split('.');
  var isOk = true;

    if (tmpVar.length < 3){
      isOk = false;
    } else {
      if (tmpVar[0] > 255) {
        isOk = false;
      } else if (tmpVar[0] < 1) {
        isOk = false;
      }

      if (tmpVar[1] > 255) {
        isOk = false;
      } else if (tmpVar[1] < 1) {
        isOk = false;
      }

      if (tmpVar[2] > 65535) {
        isOk = false;
      } else if (tmpVar[2] < 1) {
        isOk = false;
      }
    }

    return isOk;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function formConfigMSI(){
  var xmlPropertyTpl  = '';
  var isOK = true;
  var tmpAppVersion   = document.getElementsByName("AppVersion")[0].value;
  var tmpManufacturer = document.getElementsByName("Manufacturer")[0].value;
  var tmpAppName      = document.getElementsByName("AppName")[0].value;
  var tmpPKGName      = document.getElementsByName("PKGName")[0].value;
  var tmpProductCode  = document.getElementsByName("ProductCode")[0].value;
  var tmpUpgradeCode  = document.getElementsByName("UpgradeCode")[0].value;

    if (tmpAppVersion != '') {
      if (validateAppVersion(tmpAppVersionx) == false) {
        document.getElementsByName("AppVersion")[0].value = tmpAppVersion + ' is invalid version format.';
      } else {
        xmlPropertyTpl = '<Property Id="ProductVersion" Value="' + tmpAppVersion + '" />\r\n';
        isOK = false;
      }
    } else {
      document.getElementsByName("AppVersion")[0].value = 'Field is reqired.';
      isOK = false;
    }

    if (tmpManufacturer != '') {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="Manufacturer" Value="' + tmpManufacturer + '" />\r\n';
    } else {
      document.getElementsByName("Manufacturer")[0].value = 'Field is reqired.';
      isOK = false;
    }

    if (tmpAppName != '') {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ProductName" Value="' + tmpAppName + '" />\r\n';
    } else {
      document.getElementsByName("AppName")[0].value = 'Field is reqired.';
      isOK = false;
    }

    if (tmpPKGName != '') {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="PKGName" Value="' + tmpPKGName + '" />\r\n';
    } else {
      document.getElementsByName("PKGName")[0].value = 'Field is reqired.';
      isOK = false;
    }

    if (tmpProductCode != '') {
      if (uuidValidate(tmpProductCode.replace('{','').replace('}',''))) {
        xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ProductCode" Value="' + tmpProductCode + '" />\r\n';
      } else {
        document.getElementsByName("ProductCode")[0].value = tmpProductCode + ' is invalid GUID.';
        isOK = false;
      }
    } else {
      document.getElementsByName("ProductCode")[0].value = 'Field is reqired.';
      isOK = false;
    }

    if (tmpUpgradeCode != '') {
      if (uuidValidate(tmpUpgradeCode.replace('{','').replace('}',''))) {
        xmlPropertyTpl = xmlPropertyTpl + '<Property Id="UpgradeCode" Value="' + tmpUpgradeCode + '" />\r\n';
      } else {
        document.getElementsByName("UpgradeCode")[0].value = tmpUpgradeCode + ' is invalid GUID.';
        isOK = false;
      }
    } else {
      document.getElementsByName("UpgradeCode")[0].value = 'Field is reqired.';
      isOK = false;
    }

    if ($('.ui.checkbox.left.P').checkbox('is checked')) {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ARPNOMODIFY" Value="1" />\r\n';
    }

    if ($('.ui.checkbox.left.L').checkbox('is checked')) {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ARPNOREMOVE" Value="1" />\r\n';
    }

    if ($('.ui.checkbox.left.M').checkbox('is checked')) {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ARPNOREPAIR" Value="1" />\r\n';
    }


    if ($('.ui.checkbox.left.A').checkbox('is checked')) {
    }

    if ($('.ui.checkbox.left.B').checkbox('is checked')) {
    }

    if ($('.ui.checkbox.left.C').checkbox('is checked')) {
    }

    if ($('.ui.checkbox.left.D').checkbox('is checked')) {
    }


        console.log(xmlPropertyTpl);


    if (isOK === true) {

    } else {

    }
}
//----------------------------------------------------------------------------------------------------------------------
