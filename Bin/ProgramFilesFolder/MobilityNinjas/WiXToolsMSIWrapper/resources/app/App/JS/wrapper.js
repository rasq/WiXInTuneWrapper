//const prompt      = require('dialogs')(opts={});
//var Downloader    = require('mt-files-downloader');
//const homedir     = require('os').homedir();
//const mkdirp      = require('mkdirp');
//const sanitize    = require("sanitize-filename");
//const shell           = require('electron').shell;

const remote        = require('electron').remote;
const dialog        = remote.dialog;
const session       = remote.session;

const uuidV4        = require('uuid/v4');
const uuidValidate  = require('uuid-validate');

const {ipcRenderer} = require('electron');





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
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function startApp(initTab){
    $('.starter').fadeOut('fast',function(){
      $('.ui.dashboard').fadeIn('fast');
    });

    HideAll(initTab, '.ui.Opcja01');
    HideAllCheckBoxMSI();

    $('.ui.active.dimmer').fadeOut('fast');
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
    $('.ui.active.dimmer').fadeIn('fast',function(){
      formConfigMSI();
    });
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
      } else if (tmpVar[1] < 0) {
        isOk = false;
      }

      if (tmpVar[2] > 65535) {
        isOk = false;
      } else if (tmpVar[2] < 0) {
        isOk = false;
      }
    }

    return isOk;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function addCustomProp() {
    var div = document.createElement('div');

      div.className = 'row';

      div.innerHTML = '<p></p>' +
        '<div class="ui floated input" style="width: 45%;">' +
          '<input type="text" id="PropName" name="PropName" placeholder="Name">' +
        '</div>' +
        '<div class="ui right floated action input" style="margin-left: 10px; width: 45%;">' +
          '<input type="text" id="PropValue" name="PropValue" placeholder="Value">' +
          '<i class="ui green button PC" onclick="removeCustomProp(this)">-</i>' +
        '</div>';

      document.getElementById('customProperties').appendChild(div);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function removeCustomProp(input) {
    document.getElementById('customProperties').removeChild(input.parentNode.parentNode);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function addCA() {
    var div = document.createElement('div');

      div.className = 'row';

      div.innerHTML = '<p></p>' +
        '<div class="ui input" style="width: 25%;">' +
          '<input type="text" id="PropName" name="PropName" placeholder="Name">' +
        '</div>' +
        '<a class="addBF-button item" style="margin-left: 5px;">' +
          '<i id="addBF" class="ui medium green button CA" onclick="addBF()">Add File</i>' +
        '</a>' +
        '<div class="ui action input" style="margin-left: 5px; width: 25%;">' +
          '<input type="text" id="PropValue" name="PropValue" placeholder="Call function">' +
          '<i class="ui green button PC" onclick="removeCA(this)">-</i>' +
        '</div>'+
        '<div class="ui" style="display: table-cell; width: 50%;">' +
          '<select class="ui dropdown" style="width: 20%; display: table-cell;">' +
            '<option value="">Execute</option>' +
            '<option value="6">commit</option>' +
            '<option value="5">deferred</option>' +
            '<option value="4">firstSequence</option>' +
            '<option value="3">immediate</option>' +
            '<option value="2">oncePerProcess</option>' +
            '<option value="1">rollback</option>' +
            '<option value="0">secondSequence</option>' +
          '</select>' +
            '<select class="ui dropdown" style="width: 20%; margin-left: 5px; display: table-cell;">' +
              '<option value="">Impersonate</option>' +
              '<option value="1">Yes</option>' +
              '<option value="0">No</option>' +
            '</select>' +
        '</div>';




      document.getElementById('customActionDiv').appendChild(div);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function removeCA(input) {
    document.getElementById('customActionDiv').removeChild(input.parentNode.parentNode);
}
//----------------------------------------------------------------------------------------------------------------------



//----------------------------------------------------------------------------------------------------------------------
function formConfigMSI(){
  var x               = 0;
  var xmlPropertyTpl  = '';
  var isOK            = true;
  var tmpMSIConf      = new Array();
  var tmpAppVersion   = document.getElementsByName("AppVersion")[0].value;
  var tmpManufacturer = document.getElementsByName("Manufacturer")[0].value;
  var tmpAppName      = document.getElementsByName("AppName")[0].value;
  var tmpPKGName      = document.getElementsByName("PKGName")[0].value;
  var tmpProductCode  = document.getElementsByName("ProductCode")[0].value;
  var tmpUpgradeCode  = document.getElementsByName("UpgradeCode")[0].value;
  var tmpPropName     = document.getElementsByName("PropName");
  var tmpPropValue    = document.getElementsByName("PropValue");

  var tmpFiR          = "Field is reqired.";




    if (tmpAppVersion != '' && tmpAppVersion != tmpFiR) {
      if (validateAppVersion(tmpAppVersion) == false) {
        document.getElementsByName("AppVersion")[0].value = tmpAppVersion + ' is invalid version format.';
        isOK = false;
      } else {
        tmpMSIConf[1] = tmpAppVersion;
        //xmlPropertyTpl = '<Property Id="ProductVersion" Value="' + tmpAppVersion + '" />\r\n';
      }
    } else {
      document.getElementsByName("AppVersion")[0].value = tmpFiR;
      isOK = false;
    }

    if (tmpManufacturer != '' && tmpManufacturer != tmpFiR) {
      tmpMSIConf[0] = tmpManufacturer;
      //xmlPropertyTpl = xmlPropertyTpl + '<Property Id="Manufacturer" Value="' + tmpManufacturer + '" />\r\n';
    } else {
      document.getElementsByName("Manufacturer")[0].value = tmpFiR;
      isOK = false;
    }

    if (tmpAppName != '' && tmpAppName != tmpFiR) {
      tmpMSIConf[2] = tmpAppName;
      //xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ProductName" Value="' + tmpAppName + '" />\r\n';
    } else {
      document.getElementsByName("AppName")[0].value = tmpFiR;
      isOK = false;
    }

    if (tmpPKGName != '' && tmpPKGName != tmpFiR) {
      xmlPropertyTpl = xmlPropertyTpl + '<Property Id="PKGName" Value="' + tmpPKGName + '" />\r\n';
    } else {
      document.getElementsByName("PKGName")[0].value = tmpFiR;
      isOK = false;
    }

    if (tmpProductCode != '' && tmpProductCode != tmpFiR) {
      if (uuidValidate(tmpProductCode.replace('{','').replace('}',''))) {
        tmpMSIConf[3] = tmpProductCode;
        //xmlPropertyTpl = xmlPropertyTpl + '<Property Id="ProductCode" Value="' + tmpProductCode + '" />\r\n';
      } else {
        document.getElementsByName("ProductCode")[0].value = tmpProductCode + ' is invalid GUID.';
        isOK = false;
      }
    } else {
      document.getElementsByName("ProductCode")[0].value = tmpFiR;
      isOK = false;
    }

    if (tmpUpgradeCode != '' && tmpUpgradeCode != tmpFiR) {
      if (uuidValidate(tmpUpgradeCode.replace('{','').replace('}',''))) {
        tmpMSIConf[4] = tmpUpgradeCode;
        //xmlPropertyTpl = xmlPropertyTpl + '<Property Id="UpgradeCode" Value="' + tmpUpgradeCode + '" />\r\n';
      } else {
        document.getElementsByName("UpgradeCode")[0].value = tmpUpgradeCode + ' is invalid GUID.';
        isOK = false;
      }
    } else {
      document.getElementsByName("UpgradeCode")[0].value = tmpFiR;
      isOK = false;
    }




    for (x = 0; x < tmpPropName.length; x++) {
      if (tmpPropName[x].value != '' && tmpPropValue[x].value != '') {
        xmlPropertyTpl = xmlPropertyTpl + '<Property Id="' + tmpPropName[x].value + '" Value="' + tmpPropValue[x].value + '" />\r\n';
      }
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
      customActions();
    }

    if ($('.ui.checkbox.left.B').checkbox('is checked')) {
      launchConditions();
    }

    if ($('.ui.checkbox.left.C').checkbox('is checked')) {
      schortcuts();
    }

    if ($('.ui.checkbox.left.D').checkbox('is checked')) {
      registry();
    }




    if (isOK === true) {
            console.log('itsOK');

        try {
          ipcRenderer.send('send-msiConf', tmpMSIConf); //sendSync
        } catch (er) {}

        try {
          ipcRenderer.send('send-xml-properties', xmlPropertyTpl);
        } catch (er) {}

          ipcRenderer.send('send-xml-generate', '');

          ipcRenderer.send('send-xml-normalize');
              console.log('xml generated');

        ipcRenderer.send('send-wxsPath', 'path');
    } else {
            console.log('itsNotOK');
    }


      $('.ui.active.dimmer').fadeOut('fast');
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function pause(milliseconds) {
	var dt = new Date();
	 while ((new Date()) - dt <= milliseconds) { /* Do nothing */ }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function customActions() {
  var dllTPL = "DllEntry='FooEntryPoint'";
  var customActionTPL = "<?xml version='1.0'?>" +
  "<Wix xmlns='http://schemas.microsoft.com/wix/2006/wi'>" +
     '<Fragment>' +
        "<CustomAction Id='FooAction' BinaryKey='FooBinary' Execute='immediate' Return='check'/>" +
        "<Binary Id='FooBinary' SourceFile='foo.dll'/>" +
     "</Fragment>" +
  "</Wix>";



}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function launchConditions() {

}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function schortcuts() {

}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function registry() {

}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function generateMSI(arg) {
  if (arg === 0) {
    ipcRenderer.sendSync('send-msi-generate', 0);
    ipcRenderer.sendSync('send-msi-generate', 1);
    ipcRenderer.send('send-wxsobjPath', 'path');
  }

  if (arg === 2) {
    ipcRenderer.sendSync('send-msi-generate', 2);
  }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-wxsPath', (event, arg) => {
  wxsFilePathChange(arg);
})
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-wxsobjPath', (event, arg) => {
  wxsobjFilePathChange(arg);
})
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function wxsFilePathChange(path) {
  var pathDiv = document.getElementById("wxsFilePath");
    pathDiv.innerHTML = path;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function wxsobjFilePathChange(path) {
  var pathDiv = document.getElementById("wxsobjFilePath");
    pathDiv.innerHTML = path;
}
//----------------------------------------------------------------------------------------------------------------------
