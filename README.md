#VCF_Outlook_formator
This PowerShell script imports contacts from a VCF file into Microsoft Outlook.

Prerequisites
Windows operating system
Microsoft Outlook installed
PowerShell
Installation
Clone the repository to your local computer:

sh
Code kopieren
git clone https://github.com/your-username/vcf_Outlook_formator.git
Navigate to the repository directory:

sh
Code kopieren
cd vcf_Outlook_formator
Usage
Ensure that Microsoft Outlook is closed.

Open PowerShell as an administrator.

Run the script:

sh
Code kopieren
./vcf_Outlook_formator.ps1
A window will open where you can select the VCF file.

Click on "Start Import" to begin importing the contacts.

Functions
Stop-Outlook: Checks if Outlook is active and quits it if necessary.
addContacttoOutlook: Imports contacts from the selected VCF file into Outlook and displays progress in a progress bar.
Update-ProgressBar: Updates the progress bar during the import.
Troubleshooting
Error message during import: Ensure that Outlook is installed and configured.
Issues with file selection: Make sure you are selecting a valid VCF file.
Script not working as expected: Ensure you are running PowerShell with administrator rights.
