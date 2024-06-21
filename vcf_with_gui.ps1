# Lade die erforderlichen .NET-Typen
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Funktion zum Überprüfen und Beenden von Outlook
function Stop-Outlook {
    try {
        $Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        if ($null -ne $Outlook) {
            $Outlook.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Outlook) | Out-Null
            Write-Output "Outlook wurde erfolgreich geschlossen."
        }
    } catch {
        # Outlook ist nicht aktiv oder nicht gefunden
        Write-Output "Outlook ist nicht aktiv oder nicht gefunden."
    }
}

# Funktion zum Importieren der Kontakte in Outlook
function addContacttoOutlook {
    param (
        [string]$vcfFile,
        [string]$username,
        [System.Windows.Forms.ProgressBar]$progressBar
    )

    $desktop_verzeichnis = "C:\Users\$username\Desktop\Generierte_Kontakt_Datei"

    # Lese den Inhalt der Datei im UTF-8 Format
    $content = Get-Content -Path $vcfFile        
    
    # Speichere den Inhalt in der gleichen Datei, aber im Windows-1252 Format
    $content | Out-File -FilePath $vcfFile -Encoding Windows-1252

    $contactCount = 0
    $contactLines = @()

    if (-not (Test-Path $desktop_verzeichnis)) {
        New-Item -ItemType Directory -Path $desktop_verzeichnis
    }

    Get-Content -Path $vcfFile -Encoding Windows-1252 | ForEach-Object {
        if ($_ -eq "END:VCARD") {
            $contactCount++
            $zielpfad = "$desktop_verzeichnis\contact_$contactCount.vcf"
            $contactLines | Out-File -FilePath $zielpfad -Encoding Windows-1252
            $contactLines = @()
            
            try {
                # Outlook-Anwendung initialisieren
                $Outlook = New-Object -ComObject Outlook.Application
                # Kontaktobjekt erstellen
                $Contact = $Outlook.Session.OpenSharedItem($zielpfad)
                # Kontakt importieren
                $Contact.Save() 
            } catch {
                Write-Error "Fehler beim Importieren des Kontakts: $_"
            } finally {
                # Aufräumen
                if ($Contact) {
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Contact) | Out-Null
                }
                if ($Outlook) {
                    $Outlook.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
                }
            }

            # Fortschritt aktualisieren
            Update-ProgressBar -progressBar $progressBar
        } else {
            $contactLines += $_
        }
    }

    # Verzeichnis entfernen, falls es noch existiert
    if (Test-Path $desktop_verzeichnis){
        Remove-Item -Path $desktop_verzeichnis -Recurse -Force
    }

    # Erfolgsmeldung anzeigen
    [System.Windows.Forms.MessageBox]::Show("Alle Kontakte wurden erfolgreich importiert!", "Erfolg", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

    #Fomular Schließen 
    $form.Close()
}

# Funktion zum Aktualisieren der ProgressBar
function Update-ProgressBar {
    param (
        [System.Windows.Forms.ProgressBar]$progressBar
    )
    if ($progressBar.Value -le $progressBar.Maximum - $progressBar.Step) {
        $progressBar.Value += $progressBar.Step
    } else {
        $progressBar.Value = $progressBar.Maximum
    }
}

# Erstelle das Hauptfenster
$form = New-Object System.Windows.Forms.Form
$form.Text = "Contacts to Outlook"
$form.Size = New-Object System.Drawing.Size(380, 180)
$form.StartPosition = "CenterScreen"
$form.MaximizeBox = $False
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog

# Erstelle ein Label
$label = New-Object System.Windows.Forms.Label
$label.Text = "Wähle deine Datei aus:"
$label.AutoSize = $true
$label.Location = New-Object System.Drawing.Point(10, 10)
$form.Controls.Add($label)

# Erstelle eine Textbox zum Anzeigen des Dateipfads
$textbox = New-Object System.Windows.Forms.TextBox
$textbox.Size = New-Object System.Drawing.Size(250, 20)
$textbox.Location = New-Object System.Drawing.Point(10, 40)
$form.Controls.Add($textbox)

# Erstelle einen Button zum Öffnen des Datei-Auswahl Dialogs
$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Durchsuchen..."
$buttonBrowse.Location = New-Object System.Drawing.Point(270, 40)
$buttonBrowse.Add_Click({
    # Erstelle und zeige den OpenFileDialog
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Filter = "Kontaktdateien (*.vcf)|*.vcf*|Alle Dateien (*.*)|*.*"
    if ($fileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $textbox.Text = $fileDialog.FileName
    }
})
$form.Controls.Add($buttonBrowse)

# Erstelle einen Button zum Starten des Imports
$buttonImport = New-Object System.Windows.Forms.Button
$buttonImport.Text = "Start Import"
$buttonImport.Location = New-Object System.Drawing.Point(10, 70)
$buttonImport.Add_Click({
    $vcfFile = $textbox.Text
    $username = [System.Environment]::UserName
    
    # Vor dem Import überprüfen und ggf. Outlook schließen
    Check-and-Close-Outlook
    
    # Starte den Import
    addContacttoOutlook -vcfFile $vcfFile -username $username -progressBar $progressBar
})
$form.Controls.Add($buttonImport)

# Erstelle eine ProgressBar
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 100)
$progressBar.Size = New-Object System.Drawing.Size(300, 30)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Step = 10
$progressBar.Value = 0

# Füge die ProgressBar zum Formular hinzu
$form.Controls.Add($progressBar)

# Zeige das Formular
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
