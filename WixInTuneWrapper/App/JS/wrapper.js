//const prompt      = require('dialogs')(opts={});
//var Downloader    = require('mt-files-downloader');
//const homedir     = require('os').homedir();
//const mkdirp      = require('mkdirp');
//const sanitize    = require("sanitize-filename");
//const shell           = require('electron').shell;

const remote        = require('electron').remote;
const dialog        = remote.dialog;
const session       = remote.session;
const mainProcess   = remote.require('./main')

const uuidV4        = require('uuid/v4');
const uuidValidate  = require('uuid-validate');
const xmlParse      = require("xml-parse");

const {ipcRenderer} = require('electron');






var icoPath = '';
var inPath =  '';
var appCfg                = new Array();

var installExecuteSeqTPLH = "<InstallExecuteSequence>" +
         "<Custom Action='actionName' After='InstallFiles'/>" +
      "</InstallExecuteSequence>"









//----------------------------------------------------------------------------------------------------------------------
$('.wsx-button').click(function(){
  var wsxData = document.getElementById("selectwsxA");
    wsxData.click(function(){
  });
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.wixobj-button').click(function(){
  var wixobjData = document.getElementById("selectwixobjA");
    wixobjData.click(function(){
  });
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.getIco-button').click(function(){
  var dataIco = document.getElementById("selectIcoA");
    dataIco.click(function(){
    });
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
$('.start-button').click(function(){
  startApp('.Opcja01-sidebar');
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
document.getElementById('selectWorkdir').addEventListener('click', _ => {
    workdirFolderPathChange(mainProcess.selectDirectory());
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
document.getElementById('selectFolder').addEventListener('click', _ => {
    projectFolderPathChange(mainProcess.selectDirectory());
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

    ipcRenderer.send('send-projectFolderPath', 'path');
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

    if (tabName == '.ui.Opcja03') {
      ipcRenderer.send('send-selectWorkdir', 'path');
    }
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
      saveConf();
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
function addShortcut() {
    var div = document.createElement('div');

      div.className = 'row';

      div.innerHTML = '<p></p>' +
        '<div class="ui floated input" style="width: 45%;">' +
          '<input type="text" id="TargetFile" name="TargetFile" placeholder="Target File Name">' +
        '</div>' +
        '<div class="ui right floated action input" style="margin-left: 10px; width: 45%;">' +
          '<input type="text" id="TargetDir" name="TargetDir" placeholder="Target Directory">' +
          '<i class="ui green button PC" onclick="removeShortcut(this)">-</i>' +
        '</div>';

      document.getElementById('shortcut').appendChild(div);
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
function addLaunchCondition() {
    var div = document.createElement('div');

      div.className = 'row';

      div.innerHTML = '<p></p>' +
        '<div class="ui floated input" style="width: 45%;">' +
          '<input type="text" id="ConditionVal" name="ConditionVal" placeholder="Condition">' +
        '</div>' +
        '<div class="ui right floated action input" style="margin-left: 10px; width: 45%;">' +
          '<input type="text" id="DescriptionVal" name="DescriptionVal" placeholder="Description">' +
          '<i class="ui green button PC" onclick="removeLaunchCondition(this)">-</i>' +
        '</div>';

      document.getElementById('launchCondition').appendChild(div);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function removeShortcut(input) {
    document.getElementById('shortcut').removeChild(input.parentNode.parentNode);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function removeCustomProp(input) {
    document.getElementById('customProperties').removeChild(input.parentNode.parentNode);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function removeLaunchCondition(input) {
    document.getElementById('launchCondition').removeChild(input.parentNode.parentNode);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function gatherCA() {
  var tmpCAName           = document.getElementsByName("CAName")[0].value;
  var tmpCAType           = document.getElementsByName("CAType")[0].value;
  var tmpCAFName          = document.getElementsByName("CAFName")[0].value;
  var tmpCAPathCont       = document.getElementsByName("CAPathCont")[0].value;
  var data                = '';

    if (tmpCAPathCont == '' || tmpCAName == '' || tmpCAType == '' || tmpCAFName == '' || tmpCAType == 'Execute') {
      return 1;
    } else {
      data = tmpCAName + '|' + tmpCAType + '|' + tmpCAPathCont + '|' + tmpCAFName;
        tmpCAName = '';
        tmpCAType = '';
        tmpCAPathCont = '';
        tmpCAFName = '';
      return data;
    }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function addCA() {
    var div         = document.createElement('div');
    var dataToPush  = '';

      div.className = 'row';
      dataToPush = gatherCA();

        if (dataToPush != 1) {
          div.innerHTML = '<p></p>' +
            '<div class="ui fluid icon input">' +
              '<input type="text" id="CAfield" name="CAfield" value="' + dataToPush + '">' +
              '<i class="ui green button CA" onclick="removeCA(this)">-</i>' +
            '</div>';

          document.getElementById('customActionDiv').appendChild(div);
        }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function addCAFile() {
  var dataFile = document.getElementById("selectFileA");
    dataFile.click(function(){
    });
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function openWsxFile(fileInput) {
  var files = fileInput.files;

        for (var i = 0; i < files.length; i++) {
          var file = files[i];
          var path = document.getElementById("wxsFilePath");

            path.file = file;
            filePath = file.path;
            document.getElementsByName("wxsFilePath")[0].innerHTML = file.path;
        }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function openWixobjFile(fileInput) {
  var files = fileInput.files;

        for (var i = 0; i < files.length; i++) {
          var file = files[i];
          var path = document.getElementById("wxsobjFilePath");

            path.file = file;
            filePath = file.path;
            document.getElementsByName("wxsobjFilePath")[0].innerHTML = file.path;
        }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function filePathCA(fileInput) {
  var files = fileInput.files;

        for (var i = 0; i < files.length; i++) {
          var file = files[i];
          var path = document.getElementById("CAPathCont");

            path.file = file;
            filePath = file.path;
            document.getElementsByName("CAPathCont")[0].value = file.path;
        }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
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
            icoPath = file.path;
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

            ipcRenderer.send('send-xml-generate', varCheck(document.getElementById("projectFolderPath"), 1).trim());

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
function saveConf(){
  var wdPath = document.getElementById("projectsWorkdir").innerHTML;

    if (wdPath == 'Selected folder.') {

    } else {
        ipcRenderer.send('send-saveConf', wdPath);
    }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function saveConfigApp(){
  var x                   = 0;
  var varTmp              = '';
  var VarU                = '';
  var tmpAppConf          = new Array(20);
  var tmpAppVersion       = varCheck("AppVersion", 0);
  var tmpManufacturer     = varCheck("Manufacturer", 0);
  var tmpAppName          = varCheck("AppName", 0);
  var tmpPKGName          = varCheck("PKGName", 0);
  var tmpProductCode      = varCheck("ProductCode", 0);
  var tmpUpgradeCode      = varCheck("UpgradeCode", 0);
  var tmpPropName         = document.getElementsByName("PropName");
  var tmpPropValue        = document.getElementsByName("PropValue");
  var tmpARPNOMODIFY      = 0;
  var tmpARPNOREMOVE      = 0;
  var tmpARPNOREPAIR      = 0;
  var tmpConditionVal     = document.getElementsByName("ConditionVal");
  var tmpDescriptionVal   = document.getElementsByName("DescriptionVal");

  var tmpCAVal            = document.getElementsByName("CAfield");
  var tmpCANameVal        = '';
  var tmpCATypeVal        = '';
  var tmpCAFilePVal       = '';
  var tmpCAFNameVal       = '';

  var tmpTargetFileVal    = document.getElementsByName("TargetFile");
  var tmpTargetDirVal     = document.getElementsByName("TargetDir");
  var tmpInPath           = document.getElementById("projectFolderPath");
  var tmpIcoPath          = document.getElementById("thumbnil");



    if ($('.ui.checkbox.left.P').checkbox('is checked')) {
      tmpARPNOMODIFY = 1;
    }

    if ($('.ui.checkbox.left.L').checkbox('is checked')) {
      tmpARPNOREMOVE = 1;
    }

    if ($('.ui.checkbox.left.M').checkbox('is checked')) {
      tmpARPNOREPAIR = 1;
    }


        tmpAppConf[0]   = tmpPKGName;
        tmpAppConf[1]   = tmpManufacturer;
        tmpAppConf[2]   = tmpAppName;
        tmpAppConf[3]   = tmpAppVersion;
        tmpAppConf[4]   = tmpProductCode;
        tmpAppConf[5]   = tmpUpgradeCode;
        tmpAppConf[6]   = tmpARPNOMODIFY.toString();
        tmpAppConf[7]   = tmpARPNOREMOVE.toString();
        tmpAppConf[8]   = tmpARPNOREPAIR.toString();
        tmpAppConf[9]   = varCheck(tmpInPath, 1).trim();
        tmpAppConf[10]  = varCheck(tmpIcoPath, 1).trim();


        tmpAppConf[11]  = new Array();
        tmpAppConf[12]  = new Array();
        tmpAppConf[13]  = new Array();
        tmpAppConf[14]  = new Array();
        tmpAppConf[15]  = new Array();
        tmpAppConf[16]  = new Array();
        tmpAppConf[17]  = new Array();
        tmpAppConf[18]  = new Array();
        tmpAppConf[19]  = new Array();
        tmpAppConf[20]  = new Array();

    for (x = 0; x < tmpPropName.length; x++) {
      if (tmpPropName[x].value != '' && tmpPropValue[x].value != '') {
        tmpAppConf[11].push(tmpPropName[x].value);
        tmpAppConf[12].push(tmpPropValue[x].value);
      }
    }

    x = 0;
    if ($('.ui.checkbox.left.A').checkbox('is checked')) {
      for (x = 0; x < tmpCAVal.length; x++) {
        if (tmpCAVal[x].value != '') {
          VarU = tmpCAVal[x].toString();
          varTmp = VarU.split("|");

          tmpAppConf[13].push(varTmp[0]);
          tmpAppConf[14].push(varTmp[1]);
          tmpAppConf[15].push(varTmp[2]);
          tmpAppConf[16].push(varTmp[3]);
        }
      }
    }

    x = 0;
    if ($('.ui.checkbox.left.B').checkbox('is checked')) {
      for (x = 0; x < tmpConditionVal.length; x++) {
        if (tmpConditionVal[x].value != '' && tmpDescriptionVal[x].value != '') {
          tmpAppConf[17].push(tmpConditionVal[x].value);
          tmpAppConf[18].push(tmpDescriptionVal[x].value);
        }
      }
    }

    x = 0;
    if ($('.ui.checkbox.left.C').checkbox('is checked')) {
      for (x = 0; x < tmpTargetFileVal.length; x++) {
        if (tmpTargetFileVal[x].value != '' && tmpTargetDirVal[x].value != '') {
          tmpAppConf[19].push(tmpTargetFileVal[x].value);
          tmpAppConf[20].push(tmpTargetDirVal[x].value);
        }
      }
    }

    if ($('.ui.checkbox.left.D').checkbox('is checked')) {

    }

    return tmpAppConf;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function loadConfigApp(xml){
  var parsedXML = new xmlParse.DOM(xmlParse.parse(xml.toString()));
  var x = 0;

    var tmpAppVersion       = parsedXML.document.getElementsByTagName('AppVersion')[0].innerXML;
    var tmpManufacturer     = parsedXML.document.getElementsByTagName("Manufacturer")[0].innerXML;
    var tmpAppName          = parsedXML.document.getElementsByTagName("AppName")[0].innerXML;
    var tmpPKGName          = parsedXML.document.getElementsByTagName("root")[0].attributes.PKGName;
    var tmpProductCode      = parsedXML.document.getElementsByTagName("PCode")[0].innerXML;
    var tmpUpgradeCode      = parsedXML.document.getElementsByTagName("UCode")[0].innerXML;
    var tmpInPath           = parsedXML.document.getElementsByTagName("INPath")[0].innerXML;
    var tmpIcoPath          = parsedXML.document.getElementsByTagName("ICO")[0].innerXML;
    var tmpPropName         = parsedXML.document.getElementsByTagName("CustomPropName");
    var tmpPropValue        = parsedXML.document.getElementsByTagName("CustomPropValue");
    var tmpCAName           = parsedXML.document.getElementsByTagName("CAName");
    var tmpCAType           = parsedXML.document.getElementsByTagName("CAType");
    var tmpCAFile           = parsedXML.document.getElementsByTagName("CAFile");
    var tmpCAFunction       = parsedXML.document.getElementsByTagName("CAFunction");
    var tmpLCCond           = parsedXML.document.getElementsByTagName("LCCond");
    var tmpLCDesc           = parsedXML.document.getElementsByTagName("LCDesc");
    var tmpShortcutFile     = parsedXML.document.getElementsByTagName("ShortcutFile");
    var tmpShortcutDirectory= parsedXML.document.getElementsByTagName("ShortcutDirectory");





      if (parsedXML.document.getElementsByTagName("ARPNOMODIFY")[0].innerXML == '1') {
        $('.ui.checkbox.left.P').checkbox('set checked');
      } else {
        $('.ui.checkbox.left.P').checkbox('set unchecked');
      }

      if (parsedXML.document.getElementsByTagName("ARPNOREMOVE")[0].innerXML == '1') {
        $('.ui.checkbox.left.L').checkbox('set checked');
      } else {
        $('.ui.checkbox.left.L').checkbox('set unchecked');
      }

      if (parsedXML.document.getElementsByTagName("ARPNOREPAIR")[0].innerXML == '1') {
        $('.ui.checkbox.left.M').checkbox('set checked');
      } else {
        $('.ui.checkbox.left.M').checkbox('set unchecked');
      }


    document.getElementsByName("AppVersion")[0].value     = varCheck(tmpAppVersion, 2);
    document.getElementsByName("Manufacturer")[0].value   = varCheck(tmpManufacturer, 2);
    document.getElementsByName("AppName")[0].value        = varCheck(tmpAppName, 2);
    document.getElementsByName("PKGName")[0].value        = varCheck(tmpPKGName, 2);
    document.getElementsByName("ProductCode")[0].value    = varCheck(tmpProductCode, 2);
    document.getElementsByName("UpgradeCode")[0].value    = varCheck(tmpUpgradeCode, 2);
    document.getElementById("projectFolderPath").innerHTML= varCheck(tmpInPath, 2);
    document.getElementById("thumbnil").innerHTML         = varCheck(tmpIcoPath, 2);

      if (tmpPropName.length > 0 && tmpPropValue.length > 0 ) {
          for (x = 0; x < tmpPropValue.length; x ++) {
            addCustomProp();
            document.getElementsByName("PropName")[x].value = tmpPropName[x].innerXML;
            document.getElementsByName("PropValue")[x].value = tmpPropValue[x].innerXML;
          }
      }

      var customAction = '';
      if (tmpCAName.length > 0 && tmpCAType.length > 0 && tmpCAFile.length > 0 && tmpCAFunction.length > 0) {
          $('.ui.checkbox.left.A').checkbox('set checked');
          $('.CB_CA').fadeIn('fast');
          for (x = 0; x < tmpCAName.length; x ++) {
            addCA();
            customAction = tmpCAName[x].innerXML + '|' + tmpCAType[x].innerXML + '|' + tmpCAFile[x].innerXML + '|' + tmpCAFunction[x].innerXML;
            document.getElementsByName("CAfield")[x].value = customAction;
          }
      }

      if (tmpLCCond.length > 0 && tmpLCDesc.length > 0) {
          $('.ui.checkbox.left.B').checkbox('set checked');
          $('.CB_LC').fadeIn('fast');
          for (x = 0; x < tmpLCCond.length; x ++) {
            addLaunchCondition();
            document.getElementsByName("ConditionVal")[x].value = tmpLCCond[x].innerXML;
            document.getElementsByName("DescriptionVal")[x].value = tmpLCDesc[x].innerXML;
          }
      }

      if (tmpShortcutFile.length > 0 && tmpShortcutDirectory.length > 0) {
          $('.ui.checkbox.left.C').checkbox('set checked');
          $('.CB_Shortcuts').fadeIn('fast');
          for (x = 0; x < tmpShortcutFile.length; x ++) {
            addShortcut();
            document.getElementsByName("TargetFile")[x].value = tmpShortcutFile[x].innerXML;
            document.getElementsByName("TargetDir")[x].value = tmpShortcutDirectory[x].innerXML;
          }
      }





    //  $('.ui.checkbox.left.D').checkbox('set checked');
      //$('.CB_Reg').fadeIn('fast');

}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function varCheck(arg1, arg2) {
  var RC = '';
    if (arg2 == 0) {
      if (document.getElementsByName(arg1)[0].value == '') {
        RC = 'empty';
      } else {
        RC = document.getElementsByName(arg1)[0].value;
      }
    } else if (arg2 == 1) {
      if (arg1.innerHTML == '') {
        RC = 'empty';
      } else {
        RC = arg1.innerHTML;
      }
    } else if (arg2 == 2) {
      if (arg1 == 'empty') {
        RC = '';
      } else {
        RC = arg1;
      }
    }

  return RC;
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
function archivizeFiles() {
  ipcRenderer.send('archiwizeFiles', document.getElementsByName("PKGName")[0].value);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function generateMSI(arg) {
  if (arg === 0) {
    var tmpProjDir = document.getElementsByName('wxsFilePath')[0].innerHTML.trim().split('\\');
    var wsxFile = document.getElementsByName('wxsFilePath')[0].innerHTML.trim();
    var tmpProjDirPath = '', x = 0;

    for (x = 0; x < tmpProjDir.length - 1; ++x) {
      if (x == 0) {
        tmpProjDirPath = tmpProjDir[x];
      } else {
        if (tmpProjDir[x] == 'Project') {

        } else {
          tmpProjDirPath = tmpProjDirPath + '\\' + tmpProjDir[x];
        }
      }
    }

    var sendArg = new Array();
      sendArg.push(tmpProjDirPath);
      sendArg.push(wsxFile);

    ipcRenderer.sendSync('send-setTempProjDir', sendArg);

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

//----------------------------------------------------------------------------------------------------------------------
function projectFolderPathChange(path) {
  var pathDiv = document.getElementById("projectFolderPath");
    inPath = path;
    pathDiv.innerHTML = path;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function workdirFolderPathChange(path) {
  var pathDiv = document.getElementById("projectsWorkdir");
    inPath = path;
    pathDiv.innerHTML = path;
}
//----------------------------------------------------------------------------------------------------------------------








//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('getDirectory', (event, arg) => {

});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-wxsPath', (event, arg) => {
  wxsFilePathChange(arg);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-wxsobjPath', (event, arg) => {
  wxsobjFilePathChange(arg);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-projectFolderPath', (event, arg) => {
  projectFolderPathChange(arg);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-saveConfigApp', (event, arg) => {
  var RC = saveConfigApp();
    ipcRenderer.send('send-saveConfigApp', RC);
    //mainProcess.saveConfigXML(RC);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-loadConfigApp', (event, arg) => {
  loadConfigApp(arg);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcRenderer.on('get-selectWorkdir', (event, arg) => {
  workdirFolderPathChange(arg);
});
//----------------------------------------------------------------------------------------------------------------------
