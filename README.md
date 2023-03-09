![PowerShell Gallery Version](https://img.shields.io/powershellgallery/v/ModernActivedirectory) ![Language](https://img.shields.io/badge/Powershell-100.0%25-blue)  ![Bower](https://img.shields.io/bower/l/Bootstrap?style=plastic) ![Plateform](https://img.shields.io/badge/Platform-Windows-brightgreen) ![Download](https://img.shields.io/badge/Downoad%20ModerActiveDirectory-500-orange)

# ModernActiveDirectory

![Logo](Pictures/Logo.png "Logo")

New experience given an overview of Active Directory environment from a beautiful interactive HTML report
### What can i do with : 
- [ ] View key indicators
- [ ] Inventory of Active Directory
- [ ] Browse safely Active directory essential objects 
- [ ] Advanced searches in a simple way
- [ ] Support all Active directory languages
- [ ] No sensitive informations is exposed 
- [x] Work in corporate of any size :tada:

### Get New AD Look [Fast, Easy, Secure] just from one command
![portail](https://user-images.githubusercontent.com/49924401/224164475-b18b4ce6-f4b2-4f3a-8dcc-a07b9b49ddf0.gif)
#### Installation 
> #####  For all users (require admin privilege)
```Powershell
Install-Module -Name ModernActiveDirectory
```
> ##### For Current User (not require admin privilege)
```Powershell
Install-Module -Name ModernActiveDirectory -Scope Currentuser
```
#### Updates
```Powershell
Update-Module -Name ModernActiveDirectory
```
#### How to use
```Powershell
Get-ADModernReport
```
By default the number of objects listed is limited to 200 / if company have more than 200 objects by category use -illimiteddsearch
The report is saved on %appdata%\Temp current user by default / to change the directory use the parameters -Savepatch

```Powershell
Get-ADModernReport -illimitedsearch -SavePath C:\MyFolder
```
