################################################################################################
# addRRFromInformatToBingel
# gemaakt door Bob Leysen op 08/09/2023
#
# Dit script neemt verwerkt 2 excel bestanden (BingelIn.xlsx en InformatIn.xlsx) die
# samen met het script in 1 folder geplaatst worden
# Het resultaat is een nieuw excel bestand dat beide bestanden combineert zodat 
# geboortedatums en rijksregisternummers toegevoegd worden aan de unieke ID uit Bingel
# Dit bestand kan dan opgeladen worden in Bingel
#
# Om het te laten werken moet de powershell module ImportExcel op je computer staan
# Deze module is in staat om excel bestanden te bewerken zonder dat excel op je computer
# geinstalleerd is
# in een powershell window als admin (het hekje neem je niet mee want dat is commentaar)
# Install-Module -Name ImportExcel
#
# kolomkoppen (hoofdlettergevoelig) BingelIn.xlsx
# Unieke identificatie, Klas, Klasnummer, Voornaam, Naam, Geslacht
#
# kolomkoppen (hoofdlettergevoelig) InformatIn.xlsx
# Klas, Klasnummer, Voornaam, Naam, Geslacht, Geboortedatum, Rijksregisternr.
################################################################################################
# versie 1.0 08/09/2023
# Initiele versie
#
# versie 1.1 10/09/2023
# test toegevoegd of de input bestanden bestaan zodat de dialoogvensters om deze te selecteren
# niet altijd getoond moeten worden
# 
# versie 1.2 10/09/2023
# test of module geinstalleerd is;  Indien niet, installeer in de context van de gebruiker
#
# versie 1.3 11/09/2023
# validatie van de kolommen in de input bestanden
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

#Input Bingel
#test of er een bestand met de naam BingelIn.xlsx in de folder staat en gebruik dat
#toon dialoogvenster indien het bestand er niet staat
if(Test-Path -Path "$PSScriptRoot\BingelIn.xlsx" -PathType Leaf){
    $fileInBingel="$PSScriptRoot\BingelIn.xlsx"
}else{
    $fileInBingel = Get-FileName -WindowTitle "Input Bingel" -InitialDirectory $PSScriptRoot -Filter "Excel files (*.xlsx)|*.xlsx"
}

#Input Informat
#test of er een bestand met de naam InformatIn.xlsx in de folder staat en gebruik dat
#toon dialoogvenster indien het bestand er niet staat
if(Test-Path -Path "$PSScriptRoot\InformatIn.xlsx" -PathType Leaf){
    $fileInInformat="$PSScriptRoot\InformatIn.xlsx"
}else{
    $fileInInformat = Get-FileName -WindowTitle "Input Informat" -InitialDirectory $PSScriptRoot -Filter "Excel files (*.xlsx)|*.xlsx"
}


Import-Excel $fileinBingel -OutVariable Bingel
Import-Excel $fileinInformat -OutVariable Informat

#test of de nodige velden in Bingel In excel bestaan
if(!($Bingel[0].PSobject.Properties.name -match "naam")){
    write-host "Bingel bestand [$fileinBingel] bevat geen kolom met de kolomkop [naam]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Bingel[0].PSobject.Properties.name -match "voornaam")){
    write-host "Bingel bestand [$fileinBingel] bevat geen kolom met de kolomkop [voornaam]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Bingel[0].PSobject.Properties.name -match "unieke identificatie")){
    write-host "Bingel bestand [$fileinBingel] bevat geen kolom met de kolomkop [unieke identificatie]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Bingel[0].PSobject.Properties.name -match "klas")){
    write-host "Bingel bestand [$fileinBingel] bevat geen kolom met de kolomkop [klas]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"    
    exit(4)
}
if(!($Bingel[0].PSobject.Properties.name -match "klasnummer")){
    write-host "Bingel bestand [$fileinBingel] bevat geen kolom met de kolomkop [klasnummer]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}
if(!($Bingel[0].PSobject.Properties.name -match "geslacht")){
    write-host "Bingel bestand [$fileinBingel] bevat geen kolom met de kolomkop [geslacht]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}

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
if(!($Informat[0].PSobject.Properties.name -match "rijksregisternr.")){
    write-host "Informat bestand [$fileinInformat] bevat geen kolom met de kolomkop [rijksregisternr.]"
    Read-Host -Prompt "Druk op een toets om af te sluiten"
    exit(4)
}

#klasse van type Kind
class cKind {
    [string] $zoek
    [string] $uniekeIdentificatie
    [string] $klas
    [string] $klasnummer
    [string] $voornaam
    [string] $naam
    [string] $geslacht
    [string] $geboortedatum
    [string] $rijksregisternummer
}
#verzameling van alle kinderen uit Bingel in een array
$kinderen = @()  
$Bingel | foreach-object {
    $rij = New-Object -typename cKind
    $rij.zoek = ($_.Naam + $_.Voornaam).tolower()
    $rij.uniekeIdentificatie = $_."unieke identificatie"
    $rij.klas = $_.klas
    $rij.klasnummer = $_.klasnummer
    $rij.voornaam = $_.voornaam
    $rij.naam = $_.naam
    $rij.geslacht = $_.geslacht
    $kinderen += $rij
}

#sorteer lijst met kinderen op het veld zoek
$kinderen = $kinderen | Sort-Object -Property zoek

#voeg info van informat toe aan de lijst met kinderen
$Informat | foreach-object{
    $informatZoek = ($_.naam + $_.voornaam).tolower()
    $pos = [array]::indexof($kinderen.zoek,$informatZoek)
    if($pos -ge 0){
        ($kinderen[$pos]).geboortedatum = ([DateTime]::FromOADate($_.geboortedatum)).tostring("dd.MM.yyyy")
        ($kinderen[$pos]).rijksregisternummer = $_."Rijksregisternr."
    }
}


#schrijf output in excelbestand in folder met inputbestanden
$outputfolder = split-path -path $FileInBingel
$outputbestand = ($outputfolder +"\naarBingelMetRR-"+ (get-date -format("yyyy")).ToString()+".xlsx")
$kinderen |select-object -Property uniekeIdentificatie,klas,klasnummer,voornaam,naam,geslacht,geboortedatum,rijksregisternummer | Export-Excel $outputbestand
Read-Host -Prompt "Druk op een toets om af te sluiten"