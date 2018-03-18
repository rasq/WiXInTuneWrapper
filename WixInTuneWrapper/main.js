//const klaw      = require('klaw')
//const through2  = require('through2')

const electron      = require('electron')
const fsExtra       = require('fs-extra')
const klawSync      = require('klaw-sync')
const fs            = require('fs');
const path          = require('path')
const uuidV4        = require('uuid/v4');
const child         = require('child_process');
const runExec       = require('async-child-process');
const randomInt     = require('random-int');
const url           = require('url')
const xmlParse      = require("xml-parse");
const XMLWriter     = require('xml-writer');
const AdmZip        = require('adm-zip');
const DOMParser     = require('xmldom').DOMParser;
const XMLSerializer = require('xmldom').XMLSerializer
const Registry      = require('winreg');

const {ipcMain}     = require('electron');

const app           = electron.app
const Menu          = electron.Menu
const BrowserWindow = electron.BrowserWindow
const dialog        = electron.dialog;




var splitChar = "\\";

const MenuTemplate = [
  {
    label: 'File',
    submenu: [
      {
        label: 'Load conf file',
        click() {
          loadConfig();
        }
      }, {
        label: 'Save conf file',
        click() {
          mainWindow.webContents.send('get-saveConfigApp', '');
        }
      }, {
        label: 'Quit',
        accelerator: process.platform == 'darwin' ? 'Command+Q' : 'Ctrl+Q',
        click() {
          app.quit();
        }
      }
    ]
  }
]

  if (process.platform === 'darwin') {
    MenuTemplate.unshift({});
    splitChar = "/";
  }

  if (process.env.NODE_ENV !== 'production') {
    MenuTemplate.push({
      label: 'View',
      submenu : [
        {
          label: 'Toggle Dev Tools',
          click(item, focusedWindow){
            focusedWindow.toggleDevTools();
          }
        }
      ]
    });
  }







let paths, item, mainWindow, stats, filterFn;
var pathVar, WSXSection;


var projDirectory         = __dirname;
var wixToolsetPath        = path.join(__dirname, 'App/Dependencies/wix311-binaries');
var wixToolsetCandle      = path.join(wixToolsetPath, 'candle.exe')
var wixToolsetLight       = path.join(wixToolsetPath, 'light.exe')
/*var originalWixobj        = path.join(__dirname, 'wrapper.wixobj');
var newWixobj             = path.join(__dirname, 'Project/wrapper.wixobj');
var wsxFilePath           = path.join(__dirname, 'Project/wrapper.wsx');*/
var originalWixobj        = path.join(projDirectory, 'wrapper.wixobj');
var newWixobj             = path.join(projDirectory, 'WORKING/wrapper.wixobj');
var wsxFilePath           = path.join(projDirectory, 'WORKING/wrapper.wsx');

var wsxFile               = '';

var varDbg = false;
var tplDir = 'App/TPL/';
var tmpPath = "", tmpSplittedPath = "", fileCounter = 0, dirCounter = 0, featureCounter = 0, i = 0, featureNumber = 0;
var VarSchema = 450;
var VarProductCode = '';//uuidV4();
var VarUpgradeCode = '';//uuidV4();
var VarSoftwareVersion = '';
var VarManufacturer = '';
var VarSoftwareName = '';
var VarCabinetID = "CabinetID"

var VarUsedSysFolders = new Array();
var VarFilesCounter = new Array();
    VarFilesCounter[featureNumber] = new Array();

var VarScriptDir = process.cwd();
//console.dir(VarScriptDir);
var VarSplittedScriptDir = VarScriptDir.split(splitChar);

var tabString = '\t';
var installdirUsed = false;

var VarSystemFolderNameArr = new Array ('AdminToolsFolder', 'AppDataFolder', 'CommonAppDataFolder', 'CommonFiles64Folder',
                              'CommonFilesFolder', 'DesktopFolder', 'FavoritesFolder', 'FontsFolder',
                              'LocalAppDataFolder', 'MyPicturesFolder', 'NetHoodFolder', 'PersonalFolder',
                              'PrintHoodFolder', 'ProgramFiles64Folder', 'ProgramFilesFolder', 'ProgramMenuFolder',
                              'RecentFolder', 'SendToFolder', 'StartMenuFolder', 'StartupFolder', 'System16Folder',
                              'System64Folder', 'SystemFolder', 'TempFolder', 'TemplateFolder', 'WindowsFolder',
                              'WindowsVolume');

var wxsINSTALLDIR = '<Directory Id="INSTALLDIR">\r\n';

var wxsMainDirectoryOpenTemplate = '<Directory Id="TARGETDIR" Name="SourceDir">\r\n' + wxsINSTALLDIR;

var wxsSysDirectoryOpenTemplate = '<Directory Id="VarSystemFolderName">\r\n';

var wxsDirectoryOpenTemplate = '<Directory Id="TARGETDIR" Name="VarDirectoryName">\r\n';

var wxsDirectoryCloseTemplate = '</Directory>\r\n';

var wxsHeaderTemplate = '<?xml version="1.0" encoding="UTF-8"?>\r\n' +
tabString + '<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">\r\n' +
tabString + '<Product Id="VarProductCode" Codepage="1252" Language="1033" Manufacturer="VarManufacturer" Name="VarSoftwareName" UpgradeCode="VarUpgradeCode" Version="VarSoftwareVersion">\r\n' +
tabString + '<Package Comments="Contact:  Your local administrator" Description="VarSoftwareDesc" InstallerVersion="' + VarSchema + '" Keywords="Installer,MSI,Database" Languages="1033" Manufacturer="VarMSIVendor" Platform="x86" />\r\n' + wxsMainDirectoryOpenTemplate;

var wxsComponentTemplateOpen = '<Component Id="VarComponentName" Guid="{VarGUIDString}" KeyPath="yes">\r\n';

var wxsComponentTemplateClose = '</Component>\r\n';

var wxsFileTemplate = '<File Id="VarFileID" Compressed="yes" DiskId="1" Source="VarFilePathName"/>\r\n';

var wxsFeatureTemplate = '<Feature Id="DefaultFeature" Level="1">\r\n' +
'<ComponentRef Id="ApplicationFiles"/>\r\n' +
'</Feature>\r\n';

var wxsBinaryTemplate = '<Binary Id="VarBinaryID" SourceFile="VarBinaryPath" />\r\n';

var wxsLaunchConditionTemplate = '<Condition Message="This package is not intended for server installations.">MsiNTProductType=1</Condition>\r\n';

var wxsMediaTemplate = '<Media Id="' + VarCabinetID + '" Cabinet="CABName.cab" EmbedCab="yes" VolumeLabel="DISK1" />\r\n';

var wxsIonTemplate = '<Icon Id="VarIcoID" SourceFile="VarIcoPath" />\r\n';

var wxsFeatureOpenTemplate = wxsDirectoryCloseTemplate + wxsDirectoryCloseTemplate + '<Feature Id="VarFeatureID" ConfigurableDirectory="INSTALLDIR" Level="1" Title="VarFeatureTitle">\r\n';

var wxsComponentFeatureTemplate = '<ComponentRef Id="VarComponentRefID" />\r\n';

var wxsFeatureCloseTemplate = '</Feature>\r\n';

var wxsFooterTemplate = tabString + '<Media Id="1" Cabinet="Data1.cab" DiskPrompt="1" EmbedCab="yes" VolumeLabel="DISK1" />\r\n' +
        '<Property Id="DiskPrompt" Value="[1]" />\r\n' +
        '<Property Id="WIXUI_INSTALLDIR" Value="INSTALLDIR"/>\r\n' +
        '<Property Id="INSTALLDIR" Secure="yes" />\r\n wxsProperties' +
'</Product>\r\n' +
'</Wix>';

var wxsLaunchConditionTemplate = '<Condition Message="VarComment">\r\n' +
    '<![CDATA[VarCondition]]>\r\n' +
'</Condition>';

var wxsShortcutTemplate = '<DirectoryRef Id="ApplicationProgramsFolder">\r\n' +
    '<Component Id="ApplicationShortcut" Guid="PUT-GUID-HERE">\r\n' +
        '<Shortcut Id="ApplicationStartMenuShortcut"\r\n' +
                  'Name="My Application Name"\r\n' +
                  'Description="My Application Description"\r\n' +
                  'Target="[#myapplication.exe]"\r\n' +
                  'WorkingDirectory="APPLICATIONROOTDIRECTORY"/>\r\n' +
        '<RemoveFolder Id="CleanUpShortCut" Directory="ApplicationProgramsFolder" On="uninstall"/>\r\n' +
        '<RegistryValue Root="HKCU" Key="Software\Microsoft\MyApplicationName" Name="installed" Type="integer" Value="1" KeyPath="yes"/>\r\n' +
    '</Component>\r\n' +
'</DirectoryRef>';

var wxsShortcutFolderTemplate = '<Directory Id="ProgramMenuFolder">\r\n' +
    '<Directory Id="ApplicationProgramsFolder" Name="My Application Name"/>\r\n' +
'</Directory>';

var wxsProperties = '';








//----------------------------------------------------------------------------------------------------------------------
function harwestDirsFiles(varPathToScan, wxsName, dataType) {
    i = 0

        try {
          if (dataType == 'dirs') {
              paths = klawSync(varPathToScan, {nofile: true});
              //paths = klawSync(path.join(projDirectory, varPathToScan), {nofile: true});
          } else {
                filterFn = item => item.path.indexOf('.DS_Store') < 0
                paths = klawSync(varPathToScan, {nodir: true, filter: filterFn});
              //paths = klawSync(path.join(projDirectory, varPathToScan), {nodir: true, filter: filterFn});
          }
        } catch (er) {
          //console.error(er);
        }

          for (i = 0; i < paths.length; ++i) {
            try {
              saveWSXFile (paths[i].path, 'Project\\' + wxsName + '.wsx', dataType);
            } catch (er) {
              //console.error(er);
            }
          }

    return;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function saveWSXFile (contentTxt, fileName, dataType){
  var x = 0;

    if (wsxFile == '') {
      wsxFile = path.join(projDirectory, fileName);
      //wsxFile = path.join(projDirectory, fileName);
    }


    /*if (dataType == 'dirs') {
        WSXSection = generateWSXDirectorySection (contentTxt);
    } else */if (dataType == 'files') {
        if (asarFiles.length > 0) {
          for (x = 0; x < asarFiles.length; ++x){
            if (contentTxt.includes(asarFiles[x])) {
              return;
            }
          }
        }
          WSXSection = generateWSXFileSection (contentTxt);
    } else if (dataType == 'features') {
        WSXSection = generateWSXFeatureSection ();
    } else if (dataType == 'closeTag') {
        WSXSection = tmpFileTab + tabString + contentTxt;
    } else {
        WSXSection = contentTxt;
    }

        saveTextFile(projDirectory, fileName, WSXSection);

      /*  try {
            stats = fs.statSync(path.join(projDirectory, fileName));
              try {
                fs.appendFileSync(path.join(projDirectory, fileName), WSXSection);
              } catch (er) {
                  //console.error(er);
              }
        } catch (e) {
            try {
              fs.writeFileSync(path.join(projDirectory, fileName), WSXSection);
            } catch (er) {
                //console.error(er);
            }
        }*/

    return;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function loadTextFile (fileToRead){
  var dataToPass = "";

    try {
      dataToPass = fs.readFileSync(fileToRead, 'utf8');
    } catch (er) {
      //console.error("An error ocurred reading the file " + err);
    }

    return dataToPass;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
/*function generateWSXDirectorySection (contentTxt){
   var VarSplittedString = contentTxt.split(splitChar);

    var tmpString = wxsDirectoryOpenTemplate.replace("VarFilePathName", contentTxt).replace("VarGUIDString", uuidV4()).replace("VarFileID", VarSplittedString[VarSplittedString.length -1].replace(".","") + "_" + dirCounter + '_' + randomInt(999)); // + '_' + randomInt(999)

        tmpString = tmpString + wxsDirectoryCloseTemplate;

        dirCounter++;

    return tmpString;
}*/
//----------------------------------------------------------------------------------------------------------------------

var tmpFileTab = "";
var dirCount = 0;
var asarFiles = new Array();
//----------------------------------------------------------------------------------------------------------------------
function generateWSXFileSection (contentTxt){
    var strPos = 0, res = "";

      if (contentTxt.includes(".asar")) { //temp fix for .asar files that are like directories for Klaw-sync
        strPos = contentTxt.indexOf(".asar");
        res = contentTxt.substring(0, strPos + 5);
        contentTxt = res;
        asarFiles.push(contentTxt);
      }

    var VarSplittedString = contentTxt.split(splitChar);
    var tmpSplittedString = "", tmpCommponentName = "", tmpFnPath = "", tmpString = "", tmpTab = "";
    var x = 0, y = 0, z = 0;
    var isSysFolder = false;
    var skipDir = false;

        ++dirCount;

        for (x = 0; x < VarSplittedString.length - 1; ++x) {
            tmpFnPath = tmpFnPath + VarSplittedString[x];
        }

        if (tmpPath != tmpFnPath) {
            if (tmpPath != "") {
                tmpString = tmpFileTab + tabString + wxsComponentTemplateClose;
                  for (x = VarSplittedScriptDir.length + 2; x < tmpSplittedPath.length - 1; ++x) {
                      tmpTab = tmpTab + tabString;
                  }

                  tmpFileTab = tmpTab;

                  for (x = VarSplittedScriptDir.length + 2; x < tmpSplittedPath.length/* - 1*/; ++x) {
                      tmpString = tmpString + tmpTab + wxsDirectoryCloseTemplate;
                      tmpTab = tmpTab.replace(tabString, '');
                  }
                tmpTab = '';
            }

            for (x = VarSplittedScriptDir.length + 1; x < VarSplittedString.length - 1; ++x) {
                tmpTab = tmpTab + tabString;

                for (y = 0; y < VarSystemFolderNameArr.length; ++y){
                    if (VarSystemFolderNameArr[y] == VarSplittedString[x]) {
                        isSysFolder = true;
                        VarUsedSysFolders.push(VarSplittedString[x]);
                    }
                }

                if (isSysFolder == false) {
                    tmpString = tmpString + tmpTab + wxsDirectoryOpenTemplate.replace('VarDirectoryName', VarSplittedString[x]).replace("TARGETDIR", normalizeStringName (VarSplittedString[x]) + '_' + randomInt(999) + '_' + dirCount); //temp fix  + '_' + dirCount - remove after xml  ormalization
                } else {
                    for (z = 0; z < VarUsedSysFolders.length - 1; ++z) {
                      if (VarUsedSysFolders[z] == VarSplittedString[x]) {
                        skipDir = true;
                      }
                    }

                    if (skipDir == false) {
                      tmpString = tmpString + tmpTab + wxsSysDirectoryOpenTemplate.replace('VarSystemFolderName', VarSplittedString[x]);
                      if (installdirUsed == false){
                        tmpString = tmpString /*+ wxsINSTALLDIR*/;
                        installdirUsed = true;
                      }
                    }
                }
                isSysFolder = false;
                y = 0;
            }
                tmpString = tmpString + tmpTab + tabString + wxsComponentTemplateOpen.replace("VarGUIDString", uuidV4());

                tmpCommponentName = normalizeStringName (VarSplittedString[VarSplittedString.length -1].replace(/-/g,"_"));
                tmpCommponentName = tmpCommponentName + '_' + featureNumber + '_' + VarFilesCounter[featureNumber].length;

                tmpString = tmpString.replace("VarComponentName", tmpCommponentName);

                VarFilesCounter[featureNumber].push(tmpCommponentName);

            if (featureCounter == 999) {
                featureNumber++;
                featureCounter = 0;
            } else {
                featureCounter++;
            }
        }

        tmpString = tmpString + tmpFileTab + tabString + tabString + wxsFileTemplate;
        tmpString = tmpString.replace("VarFilePathName", contentTxt).replace("VarGUIDString", uuidV4()).replace("VarFileID", normalizeStringName (VarSplittedString[VarSplittedString.length -1]) + "_" + fileCounter);

        fileCounter++;

        tmpPath = tmpFnPath;
        tmpSplittedPath = VarSplittedString;

    return tmpString;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function generateWSXFeatureSection (){
    var tmpFeatureSection = "", x = 0, y = 0;

            if (tmpPath != "") {
                //tmpFeatureSection = wxsDirectoryCloseTemplate;
                  for (x = VarSplittedScriptDir.length + 1; x < tmpSplittedPath.length - 1; ++x) {
                      tmpFeatureSection = tmpFeatureSection + wxsDirectoryCloseTemplate;
                  }
            }

            for (x = 0; x < VarFilesCounter.length; ++x){
                tmpFeatureSection = tmpFeatureSection + wxsFeatureOpenTemplate.replace('VarFeatureID', "Feature_" + x + 1).replace('VarFeatureTitle', 'Complete_' + x+1);
                    for (y = 0; y < VarFilesCounter[x].length; ++y) {
                        tmpFeatureSection = tmpFeatureSection + wxsComponentFeatureTemplate.replace('VarComponentRefID', VarFilesCounter[x][y] /*+ '_' + x + '_' + y*/);
                    }
                tmpFeatureSection = tmpFeatureSection + wxsFeatureCloseTemplate;
            }

    return tmpFeatureSection;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function normalizeStringName (varStringCName){
    var tmpString = "";

        tmpString = varStringCName.replace(/-/g,"_").replace(/ /g,"_").replace(/~/g,"Q").replace(/@/g,"A").replace("+","P").replace("*","G").replace("$","D");

        if (Number.isInteger(parseInt(tmpString.charAt(0)))) {
            tmpString = "N" + tmpString;
        }
        if (tmpString.charAt(0) == "."){
            tmpString = tmpString.replace(".","_");
        }
        if (tmpString.charAt(0) == "@"){
            tmpString = tmpString.replace("@","A");
        }
        if (tmpString.charAt(0) == "_"){
            tmpString = tmpString.replace("_","U");
        }

    return tmpString;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function normalizeXml (varXml){
    //var xmlDoc = new xmlParse.DOM(xmlParse.parse(varXml));
    //var root = xmlDoc.document.getElementsByTagName("Directory")[0];

      //  //console.log(root.childNodes[0].parentNode);
      ////console.log(xmlDoc);
    var content = fs.readFileSync(varXml, "utf8");
    var doc = new DOMParser().parseFromString(content, 'text/xml');
    var str = new XMLSerializer().serializeToString(doc);
      ////console.info(doc);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function saveTextFile(dir, file, content) {
  try {
      stats = fs.statSync(path.join(dir, file));
        try {
          fs.appendFileSync(path.join(dir, file), content);
        } catch (er) {
            //console.error(er);
        }
  } catch (e) {
      try {
        fs.writeFileSync(path.join(dir, file), content);
      } catch (er) {
          //console.error(er);
      }
  }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function removeFile (filePath){
    if (fs.existsSync(filePath)) {
        try {
          fs.unlinkSync(filePath);
        } catch (er) {
            //console.log("An error ocurred updating the file");
            //console.error(er);
        }
        //console.log("File succesfully deleted");
    } else {
        //console.log("This file doesn't exist, cannot delete");
    }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function wixCandle() {
    console.log(wixToolsetCandle + ' , ' + wsxFile);
    var result = child.spawnSync(wixToolsetCandle,
                       [wsxFile]);

      if (result.status !== 0) {
        process.stderr.write(result.stderr);
        process.exit(result.status);
      } else {
        process.stdout.write(result.stdout);
        process.stderr.write(result.stderr);
      }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function wixMoveConf() {
  fsExtra.move(originalWixobj, newWixobj, { overwrite: true }, err => {
    if (err) {
      //console.error(err);
    } else {
      //console.error('file moved');
    }
  });
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function wixLight() {
  var result = child.spawnSync(wixToolsetLight,
                   [newWixobj]);
                   if (result.status !== 0) {
                     process.stderr.write(result.stderr);
                     process.exit(result.status);
                   } else {
                     process.stdout.write(result.stdout);
                     process.stderr.write(result.stderr);
                   }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function pause(milliseconds) {
	var dt = new Date();
	 while ((new Date()) - dt <= milliseconds) { /* Do nothing */ }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function createDirectory(directory) {
  if(!fs.existsSync(directory)){
      fs.mkdirSync(directory, 0766, function(err){
          if(err){
              //console.log(err);
          }
      });
  }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function archiveProject(zipName) {
  var zip = new AdmZip();
  var inD = path.join(projDirectory, 'Source');
  var wD = path.join(projDirectory, 'Project');

      zip.addLocalFolder(inD);
      zip.addLocalFolder(wD);
      zip.writeZip(path.join(outDirectory, zipName + ".zip"));
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function saveConfig(arg) {
  var xmlConfig = new XMLWriter;
  var x         = 0;
  var fileName = '';
  var xmlContent = '';

    /*for (x = 0; x < arg.length; x++) {
      //console.log("arg[" + x + "].length = " + arg[x].length);
    }*/

    xmlConfig.startDocument().startElement('root').writeAttribute('PKGName', arg[0]);
    xmlConfig.writeElement('Manufacturer', arg[1]);
    xmlConfig.writeElement('AppName', arg[2]);
    xmlConfig.writeElement('AppVersion', arg[3]);
    xmlConfig.writeElement('PCode', arg[4]);
    xmlConfig.writeElement('UCode', arg[5]);
    xmlConfig.writeElement('ARPNOMODIFY', arg[6]);
    xmlConfig.writeElement('ARPNOREMOVE', arg[7]);
    xmlConfig.writeElement('ARPNOREPAIR', arg[8]);
    xmlConfig.writeElement('INPath', arg[9]);
    xmlConfig.writeElement('ICO', arg[10]);

      x = 0;
      if (arg[11].length > 0 && arg[12].length > 0) {
          for (x = 0; x < arg[11].length; x ++) {
            xmlConfig.writeElement('CustomPropName', arg[11][x]);
            xmlConfig.writeElement('CustomPropValue', arg[12][x]);
          }
      }

      x = 0;
      if (arg[13].length > 0 && arg[14].length > 0 && arg[15].length > 0  && arg[16].length > 0) {
        for (x = 0; x < arg[11].length; x ++) {
          xmlConfig.writeElement('CAName', arg[13][x]);
          xmlConfig.writeElement('CAType', arg[14][x]);
          xmlConfig.writeElement('CAFile', arg[15][x]);
          xmlConfig.writeElement('CAFunction', arg[16][x]);
        }
      }

      x = 0;
      if (arg[17].length > 0 && arg[18].length > 0) {
        for (x = 0; x < arg[17].length; x ++) {
          xmlConfig.writeElement('LCCond', arg[17][x]);
          xmlConfig.writeElement('LCDesc', arg[18][x]);
        }
      }

      x = 0;
      if (arg[19].length > 0 && arg[20].length > 0) {
        for (x = 0; x < arg[19].length; x ++) {
          xmlConfig.writeElement('ShortcutFile', arg[19][x]);
          xmlConfig.writeElement('ShortcutDirectory', arg[20][x]);
        }
      }


    xmlConfig.endDocument();

    if (arg[0] == 'empty') {
      fileName = 'wixToolsWrapperConf.xml';
    } else {
      fileName = arg[0] + '.xml';
    }

    xmlContent = xmlConfig.toString();

  ////console.log(xmlContent);

      saveTextFile(projDirectory, fileName, xmlContent);
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function loadConfig() {
  var configFile = '';
  dialog.showOpenDialog({filters: [{name: 'xml', extensions: ['xml']}]}, function (fileNames) {
    try {
      configFile = loadTextFile (fileNames[0]);
    } catch (err) {

    }
    //
    mainWindow.webContents.send('get-loadConfigApp', configFile);
  });
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function enviroConfig(arg) {
  var RC = '';

    if (arg == 0) {
      createDirectory(path.join(projDirectory, 'Source'));
      createDirectory(path.join(projDirectory, 'Project'));
      createDirectory(path.join(projDirectory, 'Complete'));
      createDirectory(path.join(projDirectory, 'Dokumentation'));
    }

    originalWixobj        = path.join(__dirname, 'wrapper.wixobj');
    newWixobj             = path.join(projDirectory, 'Project/wrapper.wixobj');
    wsxFilePath           = path.join(projDirectory, 'Project/wrapper.wsx');
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function setRegPath(arg, arg1) {
  regKey = new Registry({
    hive: Registry.HKCU,
    key: '\\Software\\MobilityNinjas\\WixToolsWrapper\\'
  });

  regKey.set(arg, Registry.REG_SZ, arg1, function (err) {
    if (err) {
      console.log(err);
    }
  });
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function getRegPath(arg) {
  regKey = new Registry({
    hive: Registry.HKCU,
    key:  '\\Software\\MobilityNinjas\\WixToolsWrapper\\'
  });

    regKey.values(function (err, items) {
      if (err) {
        console.log(err);
      } else {
        for (var i = 0; i < items.length; ++i) {
          console.log(items[i].name);
          if (items[i].name == arg) {
            projDirectory = items[i].value;
            enviroConfig(0);
          }
        }
      }
    });
}
//----------------------------------------------------------------------------------------------------------------------













//----------------------------------------------------------------------------------------------------------------------
function createWindow () {
      mainWindow = new BrowserWindow({width: 1024, height: 680, icon: __dirname+'/App/images/build/icon.png', 'node-integration':true, resizable: false});

      mainWindow.loadURL(url.format({
          pathname: path.join(__dirname, 'index.html'),
          protocol: 'file:',
          slashes: true
      }))



      getRegPath('WorkSpace');
      //enviroConfig();


      mainWindow.on('closed', () => {
          mainWindow = null;
      })
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------/]*------------------------------------------------------------
/*exports.saveConfigXML = (arg) => {
  //console.log('saveConfigXML');
  saveConfig(arg);
}*/

exports.selectDirectory = () => {
  return dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
//Menu.setApplicationMenu(Menu.buildFromTemplate(MenuTemplate));
const menu = Menu.buildFromTemplate(MenuTemplate);
      Menu.setApplicationMenu(menu);
//}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('ready', createWindow);
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-saveConf', (event, arg) => {
  setRegPath('WorkSpace', arg);
  projDirectory = arg;
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-setTempProjDir', (event, arg) => {
  projDirectory = arg[0];
  enviroConfig(1);
  wsxFile = arg[1]
  event.returnValue = true;
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-saveConfigApp', (event, arg) => {
  saveConfig(arg);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-xml-properties', (event, arg) => {
    wxsProperties = arg;

    wxsFooterTemplate = wxsFooterTemplate.replace('wxsProperties', wxsProperties);

    event.returnValue = true;
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-msiConf', (event, arg) => {
    VarManufacturer = arg[0];
    VarSoftwareVersion = arg[1];
    VarSoftwareName = arg[2];
    VarProductCode = arg[3];
    VarUpgradeCode = arg[4];

    wxsHeaderTemplate = wxsHeaderTemplate.replace('VarProductCode', VarProductCode).replace('VarManufacturer', VarManufacturer).replace('VarSoftwareName', VarSoftwareName).replace('VarUpgradeCode', VarUpgradeCode).replace('VarSoftwareVersion', VarSoftwareVersion);
    removeFile (path.join(__dirname,  'WORKING/wrapper.wsx'));
    //removeFile (path.join(projDirectory, 'WORKING/wrapper.wsx'));

    event.returnValue = true;
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-xml-generate', (event, arg) => {
  var RC;
      console.log(arg);
      if (arg == 'empty' || arg) {

      } else {
//arg
      }
          //console.log('building xml 0');
      RC = saveWSXFile (wxsHeaderTemplate, 'Project/' + 'wrapper' + '.wsx', 'wxsHeader');
          //console.log('building xml 1');
      RC = harwestDirsFiles(arg, 'wrapper', 'files');
          //console.log('building xml 2');
      RC = saveWSXFile (wxsComponentTemplateClose, 'Project/' + 'wrapper' + '.wsx', 'closeTag');
          //console.log('building xml 3');
      RC = saveWSXFile ("", 'Project/' + 'wrapper' + '.wsx', 'features');
          //console.log('building xml 4');
      RC = saveWSXFile (wxsFooterTemplate, 'Project/' + 'wrapper' + '.wsx', 'wxsFooter');
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('archiwizeFiles', (event, arg) => {
  archiveProject(arg);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-wxsPath', (event, arg) => {
  event.sender.send('get-wxsPath', wsxFile);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-wxsobjPath', (event, arg) => {
  event.sender.send('get-wxsobjPath', newWixobj);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-selectWorkdir', (event, arg) => {
  console.log(projDirectory);
  event.sender.send('get-selectWorkdir', projDirectory);
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('get-projectFolderPathIN', (event, arg) => {

});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-projectFolderPath', (event, arg) => {
  event.sender.send('get-projectFolderPath', path.join(__dirname,  'IN'));
  //event.sender.send('get-projectFolderPath', path.join(projDirectory,  'IN'));
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-xml-normalize', (event, arg) => {
  var RC;

      //RC = normalizeXml (wsxFile);
          //console.log('normalizing xml ');
      event.returnValue = true;
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
ipcMain.on('send-msi-generate', (event, arg) => {
  var RC;

    if (arg === 0) {
      RC = wixCandle ();
          //console.log('wixCandle ' + arg.toString());
      event.returnValue = true;
    }

    if (arg === 1) {
      RC = wixMoveConf();
          //console.log('wixMoveConf ' + arg.toString());
      event.returnValue = true;
    }

    if (arg === 2) {
      RC = wixLight ();
          //console.log('wixLight ' + arg.toString());
      event.returnValue = true;
    }
});
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
/*ipcMain.on('selectFolder', (event, arg) => {
  var RC;

      RC = selectDirectory ();
          //console.log(RC);
      event.sender.send('getDirectory', RC);
});*/
//----------------------------------------------------------------------------------------------------------------------
