<?xml version="1.0" encoding="UTF-8"?>
    <Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
        <Product Id="{' + VarProductCode + '}" Codepage="1252" Language="1033" Manufacturer="VarManufacturer" Name="VarSoftwareName" UpgradeCode="{' + VarUpgradeCode + '}" Version="' + VarSoftwareVersion + '">
        <Package Comments="Contact:  Your local administrator" Description="VarSoftwareDesc" InstallerVersion="' + VarSchema + '" Keywords="Installer,MSI,Database" Languages="1033" Manufacturer="VarMSIVendor" Platform="x86" />
            <Directory Id="TARGETDIR" Name="SourceDir">