Add-Type -AssemblyName System.Windows.Forms
$username = $env:USERNAME

#get file location 
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog

$openFileDialog.Title = "Datei auswählen"
$openFileDialog.InitialDirectory = "Downloads"
$openFileDialog.Filter = "VCF Dateien (*.vcf*)|*.vcf*"

$openFileDialog.ShowDialog()

$selectedFile = $openFileDialog.FileName



$vcfFile = $selectedFile

# Lese den Inhalt der Datei im UTF-8 Format
$content = Get-Content -Path $vcfFile 

# Speichere den Inhalt in der gleichen Datei, aber im Windows-1252 Format
$content | Out-File -FilePath $vcfFile -Encoding Windows-1252


$contactCount = 0
$contactLines = @()

$verzeichnis = "C:\Users\$username\Desktop\Generierte_Kontakt_Datei"

if (-not (Test-Path $verzeichnis)) {
    New-Item -ItemType Directory -Path $verzeichnis
}
    
Get-Content -Path $vcfFile -Encoding Windows-1252 | ForEach-Object {
    if ($_ -eq "END:VCARD") {
        $contactCount++
        $zielpfad = "$verzeichnis/contact_$contactCount.vcf"
        $contactLines | Out-File -FilePath $zielpfad -Encoding Windows-1252
        $contactLines = @()
        # Outlook Application erstellen
        $Outlook = New-Object -ComObject Outlook.Application
        # Kontaktobjekt erstellen
        $Contact = $Outlook.Session.OpenSharedItem($zielpfad)
        # Kontakt importieren
        $Contact.Save() 

        # Aufräumen
        $Outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Contact) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

    } else {
        $contactLines += $_
    }
}

Remove-Item -Path $verzeichnis -Recurse -Force