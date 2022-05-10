# OfficeDeploymentGui
A helper GUI application for the Microsoft Office 365 deployment toolkit using Powershell and Windows Forms. Created using Sapien Powershell Studio

# Running the application
Download the latest [release](https://github.com/serialscriptr/OfficeDeploymentGui/releases) from this repository or download and edit/run the .ps1 (or psf file if you have an editor that supports this file type)

# To Do
- Provide all options avaliable here: https://config.office.com/deploymentsettings
- Provide more product suites
- Provide functionality for device based licensing
- Create option to import an existing configuration or download xml
- Create better progress tracking of downloads and installs and show visually via progress bar
- Detect environment and determine best install (Ex: RDS should have shared licensing)
- Refactor both Generate-InstallOfficeXML and Generate-DLOfficeXML
- Refactor to use background jobs so the gui can still be interacted with

# Completed
- Create way to check for updates of the Office Deployment GUI application and auto check on start
- Create option to export the xml created so far
