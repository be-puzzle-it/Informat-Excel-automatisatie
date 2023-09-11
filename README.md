# InformatNaarBingelGebRR
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
#
# Het bestand dat je kan opladen naar Bingel heeft volgende naam: naarInformatMetRR-2023.xlsx
