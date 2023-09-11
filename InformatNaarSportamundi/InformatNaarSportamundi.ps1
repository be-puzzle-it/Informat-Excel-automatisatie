################################################################################################
# InformatNaarSportamundi
# gemaakt door Bob Leysen op 11/09/2023
#
# Dit script neemt het inputbestand uit Informat (InformatIn.xlsx)
# Per klas wordt een excel gegenereerd die geimporteerd kan worden in Sportamundi
#
# Om het te laten werken moet de powershell module ImportExcel op je computer staan
# Deze module is in staat om excel bestanden te bewerken zonder dat excel op je computer
# geinstalleerd is
# in een powershell window als admin (het hekje neem je niet mee want dat is commentaar)
# Install-Module -Name ImportExcel
#
# kolomkoppen (hoofdlettergevoelig) InformatIn.xlsx
# Klas, Klasnummer, Voornaam, Naam, Geslacht, Geboortedatum
################################################################################################
# versie 1.0 11/09/2023
# Initiele versie
#
################################################################################################
Add-Type -AssemblyName System.Windows.Forms
function Get-FileName
{
<#
.SYNOPSIS
   Show an Open File Dialog and return the file selected by the user

.DESCRIPTION
   Show an Open File Dialog and return the file selected by the user

.PARAMETER WindowTitle
   Message Box title
   Mandatory - [String]

.PARAMETER InitialDirectory
   Initial Directory for browsing
   Mandatory - [string]

.PARAMETER Filter
   Filter to apply
   Optional - [string]

.PARAMETER AllowMultiSelect
   Allow multi file selection
   Optional - switch

 .EXAMPLE
   Get-FileName
    cmdlet Get-FileName at position 1 of the command pipeline
    Provide values for the following parameters:
    WindowTitle: My Dialog Box
    InitialDirectory: c:\temp
    C:\Temp\42258.txt

    No passthru paramater then function requires the mandatory parameters (WindowsTitle and InitialDirectory)

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp
   C:\Temp\41553.txt

   Choose only one file. All files extensions are allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect
   C:\Temp\8544.txt
   C:\Temp\42258.txt

   Choose multiple files. All files are allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "text file (*.txt) | *.txt"
   C:\Temp\AES_PASSWORD_FILE.txt

   Choose multiple files but only one specific extension (here : .txt) is allowed

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "Text files (*.txt)|*.txt| csv files (*.csv)|*.csv | log files (*.log) | *.log"
   C:\Temp\logrobo.log
   C:\Temp\mylogfile.log

   Choose multiple file with the same extension

.EXAMPLE
   Get-FileName -WindowTitle MyDialogBox -InitialDirectory c:\temp -AllowMultiSelect -Filter "selected extensions (*.txt, *.log) | *.txt;*.log"
   C:\Temp\IPAddresses.txt
   C:\Temp\log.log

   Choose multiple file with different extensions
   Nota :It's important to have no white space in the extension name if you want to show them

.EXAMPLE
 Get-Help Get-FileName -Full

.INPUTS
   System.String
   System.Management.Automation.SwitchParameter

.OUTPUTS
   System.String

.NOTESs
  Version         : 1.0
  Author          : O. FERRIERE
  Creation Date   : 11/09/2019
  Purpose/Change  : Initial development

  Based on different pages :
   mainly based on https://blog.danskingdom.com/powershell-multi-line-input-box-dialog-open-file-dialog-folder-browser-dialog-input-box-and-message-box/
   https://code.adonline.id.au/folder-file-browser-dialogues-powershell/
   https://thomasrayner.ca/open-file-dialog-box-in-powershell/
#>
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # WindowsTitle help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Message Box Title",
            Position = 0)]
        [String]$WindowTitle,

        # InitialDirectory help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Initial Directory for browsing",
            Position = 1)]
        [String]$InitialDirectory,

        # Filter help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Filter to apply",
            Position = 2)]
        [String]$Filter = "All files (*.*)|*.*",

        # AllowMultiSelect help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Allow multi files selection",
            Position = 3)]
        [Switch]$AllowMultiSelect
    )

    # Load Assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Open Class
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog

    # Define Title
    $OpenFileDialog.Title = $WindowTitle

    # Define Initial Directory
    if (-Not [String]::IsNullOrWhiteSpace($InitialDirectory))
    {
        $OpenFileDialog.InitialDirectory = $InitialDirectory
    }

    # Define Filter
    $OpenFileDialog.Filter = $Filter

    # Check If Multi-select if used
    if ($AllowMultiSelect)
    {
        $OpenFileDialog.MultiSelect = $true
    }
    $OpenFileDialog.ShowHelp = $true    # Without this line the ShowDialog() function may hang depending on system configuration and running from console vs. ISE.
    $OpenFileDialog.ShowDialog() | Out-Null
    if ($AllowMultiSelect)
    {
        return $OpenFileDialog.Filenames
    }
    else
    {
        return $OpenFileDialog.Filename
    }
}

# test of ImportExcel module geinstalleerd is en installeer ze wanneer dat niet het geval is
if(-not (Get-Module ImportExcel -ListAvailable)){
    write-host "module ImportExcel niet geinstalleerd;  Doen we nu"
    Install-Module ImportExcel -Scope CurrentUser -Force
}

#Input Informat
#test of er een bestand met de naam InformatIn.xlsx in de folder staat en gebruik dat
#toon dialoogvenster indien het bestand er niet staat
if(Test-Path -Path "$PSScriptRoot\InformatIn.xlsx" -PathType Leaf){
    $fileInInformat="$PSScriptRoot\InformatIn.xlsx"
}else{
    $fileInInformat = Get-FileName -WindowTitle "Input Informat" -InitialDirectory $PSScriptRoot -Filter "Excel files (*.xlsx)|*.xlsx"
}
Import-Excel $fileinInformat -OutVariable Informat

#Input KlassenIn
#test of er een bestand met de naam KlassenIn.xlsx in de folder staat en gebruik dat
#toon dialoogvenster indien het bestand er niet staat
if(Test-Path -Path "$PSScriptRoot\KlassenIn.xlsx" -PathType Leaf){
    $fileinKlassen="$PSScriptRoot\KlassenIn.xlsx"
}else{
    $fileinKlassen = Get-FileName -WindowTitle "Input Klassen" -InitialDirectory $PSScriptRoot -Filter "Excel files (*.xlsx)|*.xlsx"
}

Import-Excel $fileinKlassen -OutVariable Klassen

#test of de nodige velden in Bingel In excel bestaan

#test of de nodige kolomkoppen bestaan in het bestand Informat In
if(!($Informat[0].PSobject.Properties.name -match "naam")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [naam]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"    
    exit(4)
}
if(!($Informat[0].PSobject.Properties.name -match "voornaam")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [voornaam]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Informat[0].PSobject.Properties.name -match "geboortedatum")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [geboortedatum]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Informat[0].PSobject.Properties.name -match "klas")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [klas]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"    
    exit(4)
}
if(!($Informat[0].PSobject.Properties.name -match "klasnummer")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [klasnummer]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Informat[0].PSobject.Properties.name -match "geslacht")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [geslacht]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}

#klasse van type Kind
class cKind {
    [string] $UID
    [string] $Klas
    [string] $Naam
    [string] $Voornaam
    [string] $Email
    [string] $Geslacht
    [string] $GeboortedatumDag
    [string] $GeboortedatumMaand
    [string] $GeboortedatumJaar
}
#verzameling van alle kinderen uit Informat in een array
$kinderen = @()  
$Informat | foreach-object {
    $rij = New-Object -typename cKind
    $rij.UID = ""
    $rij.Klas = $_.klas
    $rij.Naam = $_.naam
    $rij.Voornaam = $_.voornaam
    $rij.Email = ""
    if($_.geslacht -eq "V"){
        $rij.geslacht = "Vrouw"
    }else{
        $rij.geslacht = "Man"
    }
    $rij.GeboortedatumDag = ([DateTime]::FromOADate($_.geboortedatum)).tostring("dd")
    $rij.GeboortedatumMaand = ([DateTime]::FromOADate($_.geboortedatum)).tostring("MM")
    $rij.GeboortedatumJaar = ([DateTime]::FromOADate($_.geboortedatum)).tostring("yyyy")
    $kinderen += $rij
}

#sorteer lijst met kinderen op het veld zoek
$kinderen = $kinderen | Sort-Object -Property Klas

$kinderen | Format-Table

#schrijf output in excelbestand in folder met inputbestanden
$outputfolder = split-path -path $FileInInformat

#ga over de klassen en maak een outputbestand per klas
$Klassen | foreach-Object{
    $huidigeKlas = $_.klas
#    write-host $_.klas
#    $kinderen | Where-Object {$_.klas -EQ $huidigeKlas} | Format-Table
#    write-host $huidigeKlas
    $outputbestand = ($outputfolder +"\"+$_.klas+"naarSportamundi-"+ (get-date -format("yyyy")).ToString()+".xlsx")
    $kinderen | Where-Object {$_.klas -EQ $huidigeKlas} |select-object -Property UID,naam,voornaam,email,geslacht,geboortedatumdag,geboortedatummaand,geboortedatumjaar | Export-Excel $outputbestand    
}

Read-Host -Prompt "Druk op een toets om af te sluiten"