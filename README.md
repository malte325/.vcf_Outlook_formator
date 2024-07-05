# VCF_Outlook_formator

This PowerShell script imports **contacts** from an single **VCF** file into Microsoft **Outlook**.

## Prerequisites
+ Windows operating system
+ Microsoft Outlook 
+ PowerShell 7

## Installation
Clone the repository to your local computer:
`git clone https://github.com/your-username/vcf_Outlook_formator.git`

Navigate to the repository directory:
`cd vcf_Outlook_formator`

## Usage
Ensure that Microsoft Outlook is **closed**.

### Run the script:

sh
`./vcf_Outlook_formator.ps1`
A window will open where you can select the VCF file.

Click on "**Start Import**" to begin importing the contacts.

## Functions
+ Stop-Outlook: Checks if Outlook is active and quits it if necessary.
+ addContacttoOutlook: Imports contacts from the selected VCF file into Outlook and displays progress in a progress bar.
+ Update-ProgressBar: Updates the progress bar during the import.

## Troubleshooting
+ Error message during import: Ensure that Outlook is installed and configured.
+ Issues with file selection: Make sure you are selecting a valid VCF file.
+ Script not working as expected: Ensure you are running PowerShell with administrator rights.
