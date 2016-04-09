#Requires -Version 3.0

<#
.SYNOPSIS

Skriptet renamer bildefiler etter en mal basert på EXIF-informasjonen i bildet om når bildet ble tatt.
.DESCRIPTION

Skriptet sjekker EXIF-informasjonen i ett eller flere bilder, og endrer filnavnet basert på denne informasjonen. Skriptet støtter også prefikser og suffikser.
.PARAMETER Sti

Sti til hvilken katalog eller hvilket bilde skriptet skal gå gjennom.
.PARAMETER Prefiks

Tekst som blir lagt til som et prefiks på det nye bildenavnet.
.PARAMETER Suffiks

Tekst som blir lagt til som et suffiks på det nye bildenavnet.
.EXAMPLE

Organiser-Bilder.ps1 -Sti IMG_0004.JPG

Endrer filnavnet basert på når bildet ble tatt.
.EXAMPLE


Organiser-Bilder.ps1 -Sti C:\Bilder -Suffiks "FRA_LINDA"

Endrer filnavnet på alle bilder som ligger under C:\Bilder (og alle undermapper) basert på når bildet ble tatt. Legger også til en suffiks som blir hetende _FRA_LINDA.
#>

Param(
    [Parameter(Mandatory=$False,Position=0)]
    [String] $Sti = '.\',
    [Parameter(Mandatory=$False,Position=1)]
    [String] $Prefiks = $null,
    [Parameter(Mandatory=$False,Position=2)]
    [String] $Suffiks = $null
)

Function Get-BildeDato($BildeSti) {

    $Bilde = $null
    $BildeTattDT = $null
    $EA = $ErrorActionPreference

    $ErrorActionPreference = "Stop"

    try {
        
        $Bilde = New-Object System.Drawing.Bitmap -ArgumentList $BildeSti -ErrorAction Stop
        
        $BildeTattRawData = $Bilde.GetPropertyItem(36867).Value 

        if ($BildeTattRawData -eq $null) {
            Write-Warning "$(Split-Path $BildeSti -Leaf) inneholder ingen EXIF-informasjon, og kan derfor ikke endres"
            return $BildeTattDT
        } else {
            $BildeTattString = [System.Text.Encoding]::Default.GetString($BildeTattRawData, 0, $BildeTattRawData.Length - 1)
            $BildeTattDT = [DateTime]::ParseExact($BildeTattString, 'yyyy:MM:dd HH:mm:ss', $null)
            
            return $BildeTattDT
        }

    } catch {
        if ($BildeTattRawData -eq $null -and $Bilde -ne $null) {
            Write-Warning "$(Split-Path $BildeSti -Leaf) inneholder ingen EXIF-informasjon, og kan derfor ikke endres"
        }

        if ($Bilde -eq $null) {
            Write-Warning "$(Split-Path $BildeSti -Leaf) kunne ikke åpnes. Er det en bildefil?"
        }

        return $BildeTattDT

    } finally {
        if ($Bilde -ne $null) {
            $Bilde.Dispose()
        }
    }

    $ErrorActionPreference = $EA
    return $BildeTattDT
}

Function Rename-Bilde($BildeSti) {
    $BildeTatt = Get-BildeDato $BildeSti

    if ($BildeTatt -ne $null) {
        $Extension = Get-ChildItem $BildeSti | Select-Object -ExpandProperty Extension
        $Katalog = Split-Path $BildeSti -Parent
        $NyttNavn = $NavnMal -f $BildeTatt.Year, $BildeTatt.Month, $BildeTatt.Day, $BildeTatt.Hour, $BildeTatt.Minute, $BildeTatt.Second
        $SuffiksString = ""
        $PrefiksString = ""

        if ($Prefiks) {
            $PrefiksString = "$Prefiks`_"
        }

        if ($Suffiks) {
            $SuffiksString = "_$Suffiks"
        }

        $Teller = 0
        $TellerTekst = ""
        $TestNavn = ""
        do {
            
            $TestNavn = "$Katalog\$PrefiksString$NyttNavn$TellerTekst$SuffiksString$Extension"
            $FinnesAllerede = Test-Path $TestNavn -PathType Any
            $Teller++
            $TellerTekst = "_{0:D3}" -f $Teller

        } while ($FinnesAllerede)

        try {
            $TmpHash = @{}
            Rename-Item $BildeSti $TestNavn
            $TmpHash.OriginalNavn = $BildeSti
            $TmpHash.NyttNavn = $TestNavn

            Write-Output (New-Object -TypeName PSObject -Property $TmpHash)
        } catch {
            Write-Error $_
        }

        
    }
}

$NavnMal = "{0:D4}_{1:D2}_{2:D2}_{3:D2}_{4:D2}_{5:D2}"

$FullSti = Resolve-Path $Sti | Select-Object -ExpandProperty Path
$ErKatalog = Test-Path $FullSti -PathType Container

$DllInfo = $null

try {
    $DllInfo = [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
} catch {
    Write-Error $_.Error
}

if ($DllInfo -ne $null) {
    Write-Verbose "Bruker DLL med sti `"$($DllInfo.Location)`" og versjon $($DllInfo.ImageRuntimeVersion)"

} else {
    Write-Error "Kunne ikke laste DLL-fil for å lese bildedata! Avslutter!"
    return
}

if ($ErKatalog) {
    
    Write-Verbose "Sti er en katalog"
    Write-Progress -Activity "Henter filliste"
    $FilListe = Get-ChildItem -Path $FullSti -Recurse -File
    Write-Progress -Activity "Henter filliste" -Completed
    
    $BildeTeller = 1
    $AntallBilder = $FilListe.Count
    $FilListe | ForEach-Object {
        Write-Progress -Activity "Endrer navn på bilder..." -Status "$($_.FullName) ($BildeTeller av $AntallBilder)" -PercentComplete (($BildeTeller / $AntallBilder)*100)
        
        Rename-Bilde $_.FullName

        $BildeTeller++

    }
    Write-Progress -Activity "Endrer navn på bilder..." -Completed

} else {
    Write-Verbose "Sti er en enkeltfil"
    Rename-Bilde $FullSti
}
