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
const BrowserWindow = electron.BrowserWindow

//const remote = require('remote'); 
//const dialog = remote.require('dialog');

let paths, item, mainWindow, stats;
var pathVar, WSXSection, fileCounter = 0, dirCounter = 0, i = 0;




var wxsHeader = '<?xml version="1.0" encoding="UTF-8"?>\r\n';

var wxsDirectoryTemplate = '<Directory Id="TARGETDIR" Name="SourceDir">' + 
'    \r\n\t<Directory Id="ProgramFilesFolder">' +
'    \r\n\t\t<Directory Id="INSTALLDIR" Name="Example">' +
'\r\n' +              
'    \r\n\t\t</Directory>' +
'    \r\n\t</Directory>' +
'\r\n</Directory>';

var wxsComponentTemplate = '' +
'\r\n<Component Id="VarComponentName" Guid="VarGUIDString">' + 
'    \r\n\t<File Id="VarFileID" Source="VarFilePathName" KeyPath="yes" Checksum="yes"/>' +
'\r\n</Component>\r\n';

var wxsFeatureTemplate = '' + 
'\r\n<Feature Id="DefaultFeature" Level="1">' +
'    \r\n\t<ComponentRef Id="ApplicationFiles"/>' +
'\r\n</Feature>\r\n';

var wxsHeader = '<?xml version="1.0"?>' +
'\r\n<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">' +
'    \r\n\t<Product Id="*" UpgradeCode="12345678-1234-1234-1234-111111111111" Name="Example Product Name" Version="0.0.1" Manufacturer="Example Company Name" Language="1033">' +
'    \r\n\t<Package InstallerVersion="200" Compressed="yes" Comments="Windows Installer Package"/>'

var wxsMediaTemplate = '    \r\n\t<Media Id="1" Cabinet="product.cab" EmbedCab="yes"/>\r\n';

var wxsFooter = '' + 
'    \r\n\t</Product>' +
'</Wix>';











function harwestDirsFiles(varPathToScan, wxsName, dataType) {
    i = 0
    
        try {
            if (dataType == 'dirs') {
                paths = klawSync(varPathToScan, {nofile: true})
            } else {
                paths = klawSync(varPathToScan, {nodir: true})
            }
        } catch (er) {
            console.error(er)
        }
        
        for (i = 0; i < paths.length; i++) {
            console.dir(paths[i].path)
            saveWSXFile (paths[i].path, 'WORKING/' + wxsName + '.xml', dataType);
        }
} 


//function harwestFiles() {
//    const files = [] 
//    klaw('\IN/Sources').on('readable', function () {
//        while ((item = this.read())) {
//            files.push(item.path);
//            saveXMLFile (item.path, 'WORKING/files.xml');
//        }
//    }).on('end', () => {
//    })
//    
//    const wrapper = [] 
//    klaw('\IN/Wrapper').on('readable', function () {
//        while ((item = this.read())) {
//            wrapper.push(item.path);
//            saveXMLFile (item.path, 'WORKING/wrapper.xml');
//        }
//    }).on('end', () => {
//    })    
//} 



function saveWSXFile (contentTxt, fileName, dataType){ 
    if (dataType == 'dirs') {
        WSXSection = generateWSXDirectorySection (contentTxt);
    } else {
        WSXSection = generateWSXFileSection (contentTxt);
    }

        try {
            stats = fs.statSync(fileName);
            fs.appendFile(fileName, WSXSection, (err) => {
                if(err){
                    alert("An error ocurred appending the file "+ err.message)
                } 
            });
        } catch (e) {
            fs.writeFile(fileName, WSXSection, (err) => {
                if(err){
                    alert("An error ocurred creating the file "+ err.message)
                }           
            });  
        }
}



function generateWSXDirectorySection (contentTxt){ 
    var VarSplittedString = contentTxt.split("/"); 
    
    var tmpString = xmlDirectoryTemplate.replace("VarFilePathName", contentTxt);
        tmpString = contentTxt.replace("VarGUIDString", uuidV4());
        tmpString = contentTxt.replace("VarFileID", VarSplittedString[VarSplittedString.length -1].replace(".","") + "_" + dirCounter);

        dirCounter++;
}

function generateWSXFileSection (contentTxt, fileName, dataType){ 
}














function createWindow () {
    mainWindow = new BrowserWindow({width: 800, height: 300})

    mainWindow.loadURL(url.format({
        pathname: path.join(__dirname, 'index.html'),
        protocol: 'file:',
        slashes: true
    }))
    
    harwestDirsFiles('\IN/Wrapper', 'wrapper', 'dirs');
    harwestDirsFiles('\IN/Sources', 'files', 'dirs');
   
    //harwestFiles();
    //mainWindow.webContents.openDevTools()

    mainWindow.on('closed', function () {
        mainWindow = null
    })
}

app.on('ready', createWindow)

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', function () {
  if (mainWindow === null) {
    createWindow()
  }
})

