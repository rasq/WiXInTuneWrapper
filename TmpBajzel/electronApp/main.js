const electron  = require('electron')
const fsExtra   = require('fs-extra')
const klaw      = require('klaw')
const klawSync  = require('klaw-sync') 
const through2  = require('through2')
const fs        = require('fs');
const path      = require('path')
const url       = require('url')
const uuidV4    = require('uuid/v4');

const app = electron.app
const Menu = electron.Menu
const BrowserWindow = electron.BrowserWindow

//const remote = require('remote'); 
//const dialog = remote.require('dialog');

let paths, item, mainWindow, stats, filterFn;
var pathVar, WSXSection; 
    
var varDbg = false;
var tplDir = 'App/TPL/';
var tmpPath = "", tmpSplittedPath = "", fileCounter = 0, dirCounter = 0, featureCounter = 0, i = 0, featureNumber = 0;
var VarSchema = 450;
var VarProductCode = uuidV4();
var VarUpgradeCode = uuidV4();
var VarSoftwareVersion = '1.0.0.0';
var VarCabinetID = "CabinetID"

var VarFilesCounter = new Array();
    VarFilesCounter[featureNumber] = new Array();

var VarScriptDir = process.cwd();
var VarSplittedScriptDir = VarScriptDir.split("/"); 

var VarSystemFolderNameArr = new Array ('AdminToolsFolder', 'AppDataFolder', 'CommonAppDataFolder', 'CommonFiles64Folder', 
                              'CommonFilesFolder', 'DesktopFolder', 'FavoritesFolder', 'FontsFolder', 
                              'LocalAppDataFolder', 'MyPicturesFolder', 'NetHoodFolder', 'PersonalFolder', 
                              'PrintHoodFolder', 'ProgramFiles64Folder', 'ProgramFilesFolder', 'ProgramMenuFolder', 
                              'RecentFolder', 'SendToFolder', 'StartMenuFolder', 'StartupFolder', 'System16Folder', 
                              'System64Folder', 'SystemFolder', 'TempFolder', 'TemplateFolder', 'WindowsFolder', 
                              'WindowsVolume');



var wxsMainDirectoryOpenTemplate = '\t<Directory Id="TARGETDIR" Name="SourceDir">\r\n'; 

var wxsSysDirectoryOpenTemplate = '\t<Directory Id="VarSystemFolderName">\r\n'; 

var wxsDirectoryOpenTemplate = '\t<Directory Id="TARGETDIR" Name="VarDirectoryName">\r\n'; 

var wxsDirectoryCloseTemplate = '\t</Directory>\r\n';

var wxsHeaderTemplate = '<?xml version="1.0" encoding="UTF-8"?>' +
'\r\n<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">' +
'    \r\n\t<Product Id="{' + VarProductCode + '}" Codepage="1252" Language="1033" Manufacturer="VarManufacturer" Name="VarSoftwareName" UpgradeCode="{' + VarUpgradeCode + '}" Version="' + VarSoftwareVersion + '">' +
'    \r\n\t<Package Comments="Contact:  Your local administrator" Description="VarSoftwareDesc" InstallerVersion="' + VarSchema + '" Keywords="Installer,MSI,Database" Languages="1033" Manufacturer="VarMSIVendor" Platform="x86" />\r\n' + wxsMainDirectoryOpenTemplate;

var wxsComponentTemplateOpen = '' +
'\t\t\t\t<Component Id="VarComponentName" Guid="{VarGUIDString}" KeyPath="yes">';

var wxsComponentTemplateClose = '\r\n\t\t\t\t</Component>\r\n';

var wxsFileTemplate = '    \r\n\t\t\t\t\t<File Id="VarFileID" Compressed="yes" DiskId="1" Source="VarFilePathName"/>';

var wxsFeatureTemplate = '' + 
'\r\n<Feature Id="DefaultFeature" Level="1">' +
'    \r\n\t<ComponentRef Id="ApplicationFiles"/>' +
'\r\n</Feature>\r\n';
 
var wxsBinaryTemplate = '   \r\n\t<<Binary Id="VarBinaryID" SourceFile="VarBinaryPath" />';

var wxsLaunchConditionTemplate = '<Condition Message="This package is not intended for server installations.">MsiNTProductType=1</Condition>';

var wxsMediaTemplate = '    \r\n\t<Media Id="' + VarCabinetID + '" Cabinet="CABName.cab" EmbedCab="yes" VolumeLabel="DISK1" />\r\n';

var wxsIonTemplate = '    \r\n\t<Icon Id="VarIcoID" SourceFile="VarIcoPath" />';

var wxsFeatureOpenTemplate = '    \r\n\r\n\t\t\t<Feature Id="VarFeatureID" ConfigurableDirectory="INSTALLDIR" Level="1" Title="VarFeatureTitle">';

var wxsComponentFeatureTemplate = '    \r\n\t\t\t\t<ComponentRef Id="VarComponentRefID" />';

var wxsFeatureCloseTemplate = '    \r\n\t\t\t</Feature>\r\n';

var wxsFooterTemplate = '' + 
'    \r\n\t</Product>' +
'</Wix>';









//----------------------------------------------------------------------------------------------------------------------
function harwestDirsFiles(varPathToScan, wxsName, dataType) {
    i = 0
    
        try {
            if (dataType == 'dirs') {
                paths = klawSync(varPathToScan, {nofile: true})
            } else {
                filterFn = item => item.path.indexOf('.DS_Store') < 0
                paths = klawSync(varPathToScan, {nodir: true, filter: filterFn})
            }
        } catch (er) {
            console.error(er)
        }
        
        for (i = 0; i < paths.length; i++) {
            if (varDbg) {console.dir(paths[i].path)};
            saveWSXFile (paths[i].path, 'WORKING/' + wxsName + '.wsx', dataType);
        }
} 
//----------------------------------------------------------------------------------------------------------------------



//----------------------------------------------------------------------------------------------------------------------
function saveWSXFile (contentTxt, fileName, dataType){ 
    if (varDbg) {console.dir("function started with fileName = " + fileName + " dataType = " + dataType + " contentTxt = " + contentTxt)};
    
    if (dataType == 'dirs') {
        WSXSection = generateWSXDirectorySection (contentTxt);
    } else if (dataType == 'files') {
        WSXSection = generateWSXFileSection (contentTxt);
    } else if (dataType == 'features') {
        WSXSection = generateWSXFeatureSection ();
    } else {
        if (varDbg) {console.dir("function started with fileName = " + fileName + " dataType = " + dataType + " contentTxt = " + contentTxt)};
        WSXSection = contentTxt;
    }
    
        try {
            stats = fs.statSync(fileName);
            fs.appendFile(fileName, WSXSection, (err) => {
                if(err){
                    alert("An error ocurred appending the file " + err.message)
                } 
            });
        } catch (e) {
            if (varDbg) {console.dir("creating file = " + fileName + " with data: " + WSXSection)};
            fs.writeFile(fileName, WSXSection, (err) => {
                if(err){
                    alert("An error ocurred creating the file " + err.message)
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
                alert("An error ocurred reading the file " + err.message)
            }   
            //console.dir('xmlFileToRead =' + xmlFileToRead + ' = ' + data);
            dataToPass = data;
        });
    
    return dataToPass;
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function generateWSXDirectorySection (contentTxt){ 
   /* var VarSplittedString = contentTxt.split("/"); 
    
        //console.dir(contentTxt);
    
    var tmpString = wxsDirectoryOpenTemplate.replace("VarFilePathName", contentTxt);
        tmpString = tmpString.replace("VarGUIDString", uuidV4());
        tmpString = tmpString.replace("VarFileID", VarSplittedString[VarSplittedString.length -1].replace(".","") + "_" + dirCounter);

        tmpString = tmpString + wxsDirectoryCloseTemplate;
        
        dirCounter++;
    
    return tmpString;*/
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
function generateWSXFileSection (contentTxt){ 
    var VarSplittedString = contentTxt.split("/"); 
    var tmpSplittedString = ""; 
    var isSysFolder = false;
    var x = 0;
    var y = 0;
    var tmpCommponentName = ""
    var tmpFnPath = "";
    var tmpString = "";
    var tmpTab = "";
    
        for (x = 0; x < VarSplittedString.length -1; x++) {
            tmpFnPath = tmpFnPath + VarSplittedString[x];
            //console.dir("tmpFnPath = " + tmpFnPath);
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
                tmpString = wxsComponentTemplateClose;
                
                for (x = VarSplittedScriptDir.length + 1; x < tmpSplittedPath.length - 1; x++) {
                    tmpString = tmpString + wxsDirectoryCloseTemplate
                }
            } 
            
            for (x = VarSplittedScriptDir.length + 1; x < VarSplittedString.length - 1; x++) {
                for (y = 0; y < VarSystemFolderNameArr.length; y++){
                    if (VarSystemFolderNameArr[y] == VarSplittedString[x]) {
                        isSysFolder = true;
                    }  
                }
                
                if (isSysFolder == false) {
                    tmpString = tmpString + wxsDirectoryOpenTemplate
                    tmpString = tmpString.replace('VarDirectoryName', VarSplittedString[x])
                    tmpString = tmpString.replace("TARGETDIR", normalizeStringName (VarSplittedString[x]));
                } else {
                    tmpString = tmpString + wxsSysDirectoryOpenTemplate
                    tmpString = tmpString.replace('VarSystemFolderName', VarSplittedString[x])
                }
                
                isSysFolder = false;
                y = 0;
            }
                 
                tmpString = tmpString + wxsComponentTemplateOpen;
                tmpString = tmpString.replace("VarGUIDString", uuidV4());
            
                tmpCommponentName = normalizeStringName (VarSplittedString[VarSplittedString.length -1].replace(/-/g,"_"));
                    
                tmpString = tmpString.replace("VarComponentName", tmpCommponentName);

                VarFilesCounter[featureNumber].push(tmpCommponentName);

            if (featureCounter == 999) {
                featureNumber++;
                featureCounter = 0;
            } else {
                featureCounter++;
            }
        } 
    
        tmpString = tmpString + wxsFileTemplate;
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
    
            //tmpFeatureSection = tmpFeatureSection + wxsDirectoryCloseTemplate;
            //tmpFeatureSection = wxsComponentTemplateClose;
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
                        tmpFeatureSection = tmpFeatureSection + wxsComponentFeatureTemplate.replace('VarComponentRefID', VarFilesCounter[x][y] + '_' + x + '_' + y);  
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
function buildMSI (){ 
    //candle.exe wrapper.wsx
    //light.exe wrapper.wixobj
}
//----------------------------------------------------------------------------------------------------------------------














//----------------------------------------------------------------------------------------------------------------------
function createWindow () {
    var RC;
    
    mainWindow = new BrowserWindow({width: 920, height: 680, icon: __dirname+'/App/images/build/icon.png', 'node-integration':true, resizable: false});

    
    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file:',
        slashes: true
    }))
    
    
            RC = saveWSXFile (wxsHeaderTemplate, 'WORKING/' + 'wrapper' + '.wsx', 'wxsHeader');
            
            harwestDirsFiles('\IN', 'wrapper', 'files');

            RC = saveWSXFile (wxsComponentTemplateClose, 'WORKING/' + 'wrapper' + '.wsx', 'closeTag');
    
            RC = saveWSXFile ("", 'WORKING/' + 'wrapper' + '.wsx', 'features');
   
            RC = saveWSXFile (wxsFooterTemplate, 'WORKING/' + 'wrapper' + '.wsx', 'wxsFooter');
      
    
    if (varDbg) { mainWindow.webContents.openDevTools() }
    mainWindow.webContents.openDevTools() 
    
    mainWindow.on('closed', () => {
        mainWindow = null
    })
    
    const template = [
    {
        label: 'Edit',
        submenu: [
        {role: 'undo'},
        {role: 'redo'},
        {type: 'separator'},
        {role: 'cut'},
        {role: 'copy'},
        {role: 'paste'},
        {role: 'pasteandmatchstyle'},
        {role: 'delete'},
        {role: 'selectall'}
        ]
    }]
    
    if (process.platform === 'darwin') {
        template.unshift({
            label: app.getName(),
            submenu: [
            {role: 'about'},
            {type: 'separator'},
            {role: 'services', submenu: []},
            {type: 'separator'},
            {role: 'hide'},
            {role: 'hideothers'},
            {role: 'unhide'},
            {type: 'separator'},
            {role: 'quit'}
            ]
        })

        template[1].submenu.push(
            {type: 'separator'},
            {
                label: 'Speech',
                submenu: [
                    {role: 'startspeaking'},
                    {role: 'stopspeaking'}
                ]
            }
        )
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------  
Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('ready', createWindow)
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})
//----------------------------------------------------------------------------------------------------------------------

//----------------------------------------------------------------------------------------------------------------------
app.on('activate', function () {
  if (mainWindow === null) {
    createWindow()
  }
})
//----------------------------------------------------------------------------------------------------------------------

