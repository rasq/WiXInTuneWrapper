
//const xlTpl = require('random-int');

const electron  = require('electron')
const fsExtra   = require('fs-extra')
const klaw      = require('klaw')
const klawSync  = require('klaw-sync')
const through2  = require('through2')
const fs        = require('fs');
const path      = require('path')
const url       = require('url')
const uuidV4    = require('uuid/v4');
const child     = require('child_process');
const runExec   = require('async-child-process');
const randomInt = require('random-int');

const app = electron.app
const Menu = electron.Menu
const BrowserWindow = electron.BrowserWindow

var splitChar = "\\";

const MenuTemplate = [
  {
    label: 'File',
    submenu: [
      {
        label: 'Load conf file',
        click() {}
      },
      {
        label: 'Save conf file',
        click() {}
      },
      {
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
//const remote = require('remote');
//const dialog = remote.require('dialog');

let paths, item, mainWindow, stats, filterFn;
var pathVar, WSXSection;


var wixToolsetPath        = path.join(__dirname, 'App/Dependencies/wix311-binaries');
var wixToolsetCandle      = path.join(wixToolsetPath, 'candle.exe')
var wixToolsetLight       = path.join(wixToolsetPath, 'light.exe')
var wsxFile               = '';
var wixToolsetParameters  = [""];

var varDbg = false;
var tplDir = 'App/TPL/';
var tmpPath = "", tmpSplittedPath = "", fileCounter = 0, dirCounter = 0, featureCounter = 0, i = 0, featureNumber = 0;
var VarSchema = 450;
var VarProductCode = uuidV4();
var VarUpgradeCode = uuidV4();
var VarSoftwareVersion = '1.0.0.0';
var VarCabinetID = "CabinetID"

var VarUsedSysFolders = new Array();
var VarFilesCounter = new Array();
    VarFilesCounter[featureNumber] = new Array();

var VarScriptDir = process.cwd();
console.dir(VarScriptDir);
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

var wxsMainDirectoryOpenTemplate = '<Directory Id="TARGETDIR" Name="SourceDir">\r\n';

var wxsSysDirectoryOpenTemplate = '<Directory Id="VarSystemFolderName">\r\n';

var wxsDirectoryOpenTemplate = '<Directory Id="TARGETDIR" Name="VarDirectoryName">\r\n';

var wxsDirectoryCloseTemplate = '</Directory>\r\n';

var wxsHeaderTemplate = '<?xml version="1.0" encoding="UTF-8"?>\r\n' +
tabString + '<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">\r\n' +
tabString + '<Product Id="{' + VarProductCode + '}" Codepage="1252" Language="1033" Manufacturer="VarManufacturer" Name="VarSoftwareName" UpgradeCode="{' + VarUpgradeCode + '}" Version="' + VarSoftwareVersion + '">\r\n' +
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

var wxsFeatureOpenTemplate = wxsDirectoryCloseTemplate + '<Feature Id="VarFeatureID" ConfigurableDirectory="INSTALLDIR" Level="1" Title="VarFeatureTitle">\r\n';

var wxsComponentFeatureTemplate = '<ComponentRef Id="VarComponentRefID" />\r\n';

var wxsFeatureCloseTemplate = '</Feature>\r\n';

var wxsFooterTemplate = tabString + '<Media Id="1" Cabinet="Data1.cab" DiskPrompt="1" EmbedCab="yes" VolumeLabel="DISK1" />\r\n' +
        '<Property Id="DiskPrompt" Value="[1]" />\r\n' +
        '<Property Id="INSTALLDIR" Secure="yes" />\r\n' +
'</Product>\r\n' +
'</Wix>';









//----------------------------------------------------------------------------------------------------------------------
function harwestDirsFiles(varPathToScan, wxsName, dataType) {
    i = 0

        try {
            if (dataType == 'dirs') {
                paths = klawSync(varPathToScan, {nofile: true});
                console.dir('path ' + paths[i].path);
            } else {
                filterFn = item => item.path.indexOf('.DS_Store') < 0
                paths = klawSync(varPathToScan, {nodir: true, filter: filterFn});
            }
        } catch (er) {
            console.error(er);
        }

        for (i = 0; i < paths.length; i++) {
            saveWSXFile (paths[i].path, 'WORKING/' + wxsName + '.wsx', dataType);
        }

    return;
}
//----------------------------------------------------------------------------------------------------------------------



//----------------------------------------------------------------------------------------------------------------------
function saveWSXFile (contentTxt, fileName, dataType){
    if (wsxFile == '') {
      wsxFile = path.join(__dirname, fileName);
    }

    if (varDbg) {console.dir("function started with fileName = " + fileName + " dataType = " + dataType + " contentTxt = " + contentTxt)};

    if (dataType == 'dirs') {
      WSXSection = generateWSXDirectorySection (contentTxt);
    } else if (dataType == 'files') {
        WSXSection = generateWSXFileSection (contentTxt);
    } else if (dataType == 'features') {
        WSXSection = generateWSXFeatureSection ();
    } else if (dataType == 'closeTag') {
        WSXSection = tmpFileTab + tabString + contentTxt;
    } else {
        if (varDbg) {console.dir("function started with fileName = " + fileName + " dataType = " + dataType + " contentTxt = " + contentTxt)};
        WSXSection = contentTxt;
    }

        try {
            stats = fs.statSync(fileName);
            fs.appendFile(fileName, WSXSection, (err) => {
                if(err){
                    alert("An error ocurred appending the file " + err.message);
                }
            });
        } catch (e) {
            if (varDbg) {console.dir("creating file = " + fileName + " with data: " + WSXSection)};
            fs.writeFile(fileName, WSXSection, (err) => {
                if(err){
                    alert("An error ocurred creating the file " + err.message);
                }
            });
        }

    return 0;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function loadXMLTemplateFile (xmlFileToRead){
    var dataToPass = "";

        fs.readFile(xmlFileToRead, 'utf8', (err, data) => {
            if (err) {
                alert("An error ocurred reading the file " + err.message);
            }
            dataToPass = data;
        });

    return dataToPass;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function generateWSXDirectorySection (contentTxt){
   var VarSplittedString = contentTxt.split(splitChar);

        //console.dir(contentTxt);

    var tmpString = wxsDirectoryOpenTemplate.replace("VarFilePathName", contentTxt);
        tmpString = tmpString.replace("VarGUIDString", uuidV4());
        tmpString = tmpString.replace("VarFileID", VarSplittedString[VarSplittedString.length -1].replace(".","") + "_" + dirCounter);

        tmpString = tmpString + wxsDirectoryCloseTemplate;

        dirCounter++;

    return tmpString;
}
//----------------------------------------------------------------------------------------------------------------------

var tmpFileTab = "";
//----------------------------------------------------------------------------------------------------------------------
function generateWSXFileSection (contentTxt){
    var VarSplittedString = contentTxt.split(splitChar);
    var tmpSplittedString = "";
    var isSysFolder = false;
    var skipDir = false;
    var x = 0;
    var y = 0;
    var z = 0;
    var tmpCommponentName = ""
    var tmpFnPath = "";
    var tmpString = "";
    var tmpTab = "";

        for (x = 0; x < VarSplittedString.length -1; x++) {
            tmpFnPath = tmpFnPath + VarSplittedString[x];
        }

        if (varDbg) {
            console.dir("---------------------------------------------------");
            console.dir("tmpPath = " + tmpPath);
            console.dir("tmpFnPath = " + tmpFnPath);
            console.dir("---------------------------------------------------");
            console.dir(" ");
        }

        if (tmpPath != tmpFnPath) {
            if (tmpPath != "") {
                tmpString = tmpFileTab + tabString + wxsComponentTemplateClose;

                for (x = VarSplittedScriptDir.length + 2; x < tmpSplittedPath.length - 1; x++) {
                    tmpTab = tmpTab + tabString;
                }
                tmpFileTab = tmpTab;

                for (x = VarSplittedScriptDir.length + 2; x < tmpSplittedPath.length - 1; x++) {
                    tmpString = tmpString + tmpTab + wxsDirectoryCloseTemplate;
                    tmpTab = tmpTab.replace(tabString, '');
                }
                tmpTab = '';
            }

            for (x = VarSplittedScriptDir.length + 1; x < VarSplittedString.length - 1; x++) {
                tmpTab = tmpTab + tabString;

                for (y = 0; y < VarSystemFolderNameArr.length; y++){
                    if (VarSystemFolderNameArr[y] == VarSplittedString[x]) {
                        isSysFolder = true;
                        VarUsedSysFolders.push(VarSplittedString[x]);
                    }
                }

                if (isSysFolder == false) {
                    tmpString = tmpString + tmpTab + wxsDirectoryOpenTemplate;
                    tmpString = tmpString.replace('VarDirectoryName', VarSplittedString[x]);
                    tmpString = tmpString.replace("TARGETDIR", normalizeStringName (VarSplittedString[x]) + '_' + randomInt(999));
                } else {
                    for (z = 0; z < VarUsedSysFolders.length - 1; z++) {
                      if (VarUsedSysFolders[z] == VarSplittedString[x]) {
                        skipDir = true;
                      }
                    }

                    if (skipDir == false) {
                      tmpString = tmpString + tmpTab + wxsSysDirectoryOpenTemplate;
                      tmpString = tmpString.replace('VarSystemFolderName', VarSplittedString[x]);
                      if (installdirUsed == false){
                        tmpString = tmpString + wxsINSTALLDIR;
                        installdirUsed = true;
                      }
                    } else {
                      console.dir("Skipping dir, used before " + VarSplittedString[x]);
                    }
                }

                isSysFolder = false;
                y = 0;
            }

                tmpString = tmpString + tmpTab + tabString + wxsComponentTemplateOpen;
                tmpString = tmpString.replace("VarGUIDString", uuidV4());

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
        tmpString = tmpString.replace("VarFilePathName", contentTxt);
        tmpString = tmpString.replace("VarGUIDString", uuidV4());
        tmpString = tmpString.replace("VarFileID", normalizeStringName (VarSplittedString[VarSplittedString.length -1]) + "_" + fileCounter);

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
                tmpFeatureSection = wxsDirectoryCloseTemplate;
                for (x = VarSplittedScriptDir.length + 1; x < tmpSplittedPath.length - 1; x++) {
                    tmpFeatureSection = tmpFeatureSection + wxsDirectoryCloseTemplate;
                }
            }

            for (x = 0; x < VarFilesCounter.length; x++){
                tmpFeatureSection = tmpFeatureSection + wxsFeatureOpenTemplate.replace('VarFeatureID', "Feature_" + x + 1);
                tmpFeatureSection = tmpFeatureSection.replace('VarFeatureTitle', 'Complete_' + x+1);
                    for (y = 0; y < VarFilesCounter[x].length; y++) {
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

        tmpString = varStringCName.replace(/-/g,"_");
        tmpString = tmpString.replace(/ /g,"_");
        tmpString = tmpString.replace("+","P");
        tmpString = tmpString.replace("*","G");
        tmpString = tmpString.replace("$","D");

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
function removeFile (filePath){
    if (fs.existsSync(filePath)) {
        fs.unlink(filePath, (err) => {
            if (err) {
                console.log("An error ocurred updating the file" + err.message);
                console.log(err);
                return;
            }
            console.log("File succesfully deleted");
        });
    } else {
        console.log("This file doesn't exist, cannot delete");
    }
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function buildMSI (){
  var originalWixobj = path.join(__dirname, 'wrapper.wixobj');
  var newWixobj = path.join(__dirname, 'WORKING/wrapper.wixobj');
  var resolve = '';
    wixToolsetParameters = [wsxFile];

    resolve = wixToolsetCandle + ' ' + '"' + wsxFile + '"';
    /*var wixToolsetPath        = path.join(__dirname, 'App/Dependencies/wix311-binaries');
    var wixToolsetCandle      = path.join(wixToolsetPath, 'candle.exe')
    var wixToolsetLight       = path.join(wixToolsetPath, 'light.exe')*/


    const cmd = wixToolsetCandle;

    var result = child.spawnSync(cmd,
                       [path.join(wixToolsetPath, 'wrapper.wsx')]);

    if (result.status !== 0) {
      process.stderr.write(result.stderr);
      process.exit(result.status);
    } else {
      process.stdout.write(result.stdout);
      process.stderr.write(result.stderr);

        fsExtra.move(originalWixobj, newWixobj, { overwrite: true }, err => {
          if (err) {
            console.error(err);
          } else {
            var resultB = child.spawnSync(wixToolsetLight,
                               [newWixobj]);
                               if (result.status !== 0) {
                                 process.stderr.write(result.stderr);
                                 process.exit(result.status);
                               } else {
                                 process.stdout.write(result.stdout);
                                 process.stderr.write(result.stderr);
                               }
          }
        });
    }






    //var result = child.execSync('candle.exe', wixToolsetPath).toString();
      //console.log(result);

    //

    //var testCode = child.spawnSync(cmd, ['-?']);


    /*var testCode = child.execSync(cmd,
      {
        cwd: __dirname,
        input: wsxFile
      }
    );*/

    //console.log(testCode.toString());


              /*fsExtra.move(originalWixobj, newWixobj, { overwrite: true }, err => {
                if (err) {
                  console.error(err);
                  console.log(data.toString());
                } else {
                  console.log(data.toString());
                }
              })*/

    //resolve = wixToolsetLight + " " + newWixobj;

    //child.execSync(resolve);

    /*child(wixToolsetCandle, wixToolsetParameters, function(err, data) {
        if(err) {
          console.log(err);
          console.log(data.toString());
        } else {
          console.log(originalWixobj + ' created.');
              fsExtra.move(originalWixobj, newWixobj, { overwrite: true }, err => {
                if (err) {
                  console.error(err)
                } else {
                  console.log('success!')
                }
              })
        }
    });

    wixToolsetParameters = [newWixobj];


runExec
    child(wixToolsetLight, wixToolsetParameters, function(err, data) {
      if(err) {
        console.log(err);
        console.log(data.toString());
      } else {
        console.log(data.toString());
      }
    });*/
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function buildWSX (){
    var RC;

    RC = saveWSXFile (wxsHeaderTemplate, 'WORKING/' + 'wrapper' + '.wsx', 'wxsHeader');
    RC = harwestDirsFiles('\IN', 'wrapper', 'files');
    RC = saveWSXFile (wxsComponentTemplateClose, 'WORKING/' + 'wrapper' + '.wsx', 'closeTag');
    RC = saveWSXFile ("", 'WORKING/' + 'wrapper' + '.wsx', 'features');
    RC = saveWSXFile (wxsFooterTemplate, 'WORKING/' + 'wrapper' + '.wsx', 'wxsFooter');

      //pause(3000);

          /*fsExtra.copy(wsxFile, path.join(wixToolsetPath, 'wrapper.wsx'), function (err) {
            if (err) {
              console.error(err);
            }
          });*/

      //buildMSI ();
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function pause(milliseconds) {
	var dt = new Date();
	while ((new Date()) - dt <= milliseconds) { /* Do nothing */ }
}
//----------------------------------------------------------------------------------------------------------------------












//----------------------------------------------------------------------------------------------------------------------
function createWindow () {
      mainWindow = new BrowserWindow({width: 920, height: 680, icon: __dirname+'/App/images/build/icon.png', 'node-integration':true, resizable: false});

      mainWindow.loadURL(url.format({
          pathname: path.join(__dirname, 'index.html'),
          protocol: 'file:',
          slashes: true
      }))

        removeFile (path.join(__dirname,  'WORKING/wrapper.wsx'));
        buildWSX ();

      mainWindow.on('closed', () => {
          mainWindow = null;
      })
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
//Menu.setApplicationMenu(Menu.buildFromTemplate(MenuTemplate));
const menu = Menu.buildFromTemplate(MenuTemplate);
      Menu.setApplicationMenu(menu);
//}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('ready', createWindow)
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit();
  }
})
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('activate', function () {
  if (mainWindow === null) {
    createWindow();
  }
})
//----------------------------------------------------------------------------------------------------------------------
