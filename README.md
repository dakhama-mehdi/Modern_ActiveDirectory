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
- [ ] Faster report building
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
#### Usage
```Powershell
Get-ADModernReport
```
- [ ]  By default the number of objects listed is limited to 200 / if company have more than 200 objects by category use -illimiteddsearch
- [ ]  The report is saved on %appdata%\Temp current user by default / to change the directory use the parameters -Savepatch

```Powershell
Get-ADModernReport -illimitedsearch -SavePath C:\MyFolder
```
### Dependencies
- [x] RSAT if run externally from Windows 10 machine
- [ ] Powershell 5.1 or more
- [ ] PSWriteHTML Module (if use offline install)
### Variables
| parameters  | Description |
| ------------- | ------------- |
| CompanyLogo   | Logo that will be in the upper left corner of the report  |
| Content Cell  | Content Cell  |
| RightLogo     | The logo that will be in the upper right corner of the report |
| ReportTitle   | the title of the report |
| SavePath      | where the report will be saved (Example : C:\report ) |
| Days          | Defines the days for "Search for users who have not logged in for X days" |
| UserCreatedDays | Defines the days to "Get users who have been created in X days or less" |
| DayUntilPWExpireINT | Sets the days to "Get users whose passwords expire in less than X days" |
| Maxsearcher | Maximum number of Computer/User objects to search |
| OUlevelSearch | Search level in OUs (Base/Onelevel/Subtree) |
| IllimitedSearch | Search in all objects without number limits |
| Showadmin | Display the administrators in the result |
| HtmlOnePage | Generates a report in one page, (recommended for small companies) |

## Credits
### MVP Members 
- [Przemyslaw Klys](https://www.linkedin.com/in/pklys/) author of PSWriteHTML - without him this wouldn't be possible [Github](https://github.com/EvotecIT/PSWriteHTML)
- [Brad Wyatt](https://www.thelazyadministrator.com/) author of inspired project - without him this wouldn't be possible [Github](https://github.com/bwya77)
- [Thirrey DEMAN-BARCELO](https://www.experts-exchange.com/members/DEMAN-BARCELOMVP-Thierry.html)
- [Florian Burnel](https://www.it-connect.fr/author/florian/)
### Other members
Matthiew Souin, Hatira Mahmoud, Sarouti Zouhair
### Thanks 
![Credits](Pictures/Credits1.png "Credits")
