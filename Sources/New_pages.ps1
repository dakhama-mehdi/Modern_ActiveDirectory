New-HTML -TitleText 'AD_OVH' -ShowHTML -Online -FilePath $ReportSavePath {
   
    New-HTMLNavTop -Logo $CompanyLogo -MenuColorBackground gray  -MenuColor Black -HomeColorBackground gray  -HomeLinkHome   {
       
        New-NavTopMenu -Name 'Domains' -IconRegular address-book -IconColor black  {
        New-NavLink -IconSolid users -Name 'Groups' -InternalPageID 'Groups'
        New-NavLink -IconMaterial folder -Name 'OU' -InternalPageID 'OU'
        New-NavLink -IconSolid scroll -Name 'Group Policy' -InternalPageID 'GPO'
        }

        New-NavTopMenu -Name 'Objects' -IconSolid sitemap {
            New-NavLink -IconSolid user-tie -Name 'Users' -InternalPageID 'Users'
            New-NavLink -IconSolid laptop -Name 'Computers' -InternalPageID 'Computers'
        }

        New-NavTopMenu -Name 'About' -IconRegular chart-bar {
            New-NavLink -IconSolid chart-pie -Name 'Resume' -InternalPageID 'Resume'
        }
    } 
   
    New-HTMLTab -Name 'Dashboard' -IconRegular chart-bar  {
    New-HTMLTabStyle  -BackgroundColorActive teal
        
       New-HTMLSection -Name 'Company Information' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons -DataTable $CompanyInfoTable -DisablePaging -DisableSelect -DisableSearch -DisableStateSave -DisableInfo 
            }
        }    
       Section -Name 'Groups' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            Section -Name 'Domain Administrators' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $DomainAdminTable
            }
            Section -Name 'Enterprise Administrators' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $EnterpriseAdminTable
            }

        }
       Section -Name 'Objects in Default OUs' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            Section -Name 'Computers' -HeaderBackGroundColor teal  {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -HideFooter -DataTable $DefaultComputersinDefaultOUTable 
            }
            Section -Name 'Users' -HeaderBackGroundColor teal {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -HideFooter -DataTable $DefaultUsersinDefaultOUTable
            }

        }   
             
       Section -Name 'AD Objects Deleted in Last 5 Days' -HeaderBackGroundColor teal -HeaderTextAlignment left {
               Panel {
                new-htmlTable -HideFooter -DataTable $ADObjectTable
            }

        }
       Section -Name 'Expiring Items' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            Section -Name "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $PasswordExpireSoonTable 
            }
            Section -Name 'Accounts Expiring Soon' -HeaderBackGroundColor teal {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -HideFooter -DataTable $ExpiringAccountsTable
            }

        }

       Section -Name 'Accounts' -HeaderBackGroundColor teal -HeaderTextAlignment left  {

            Section -Name "Users Haven't Logged on in $Days Days or more" -HeaderBackGroundColor teal  {
                new-htmlTable -HideFooter -DataTable $userphaventloggedonrecentlytable  
            }
            Section -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $NewCreatedUsersTable
            }

        }
       Section -Name 'Security Logs' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         Panel {
                new-htmlTable -HideFooter -HideButtons -DataTable $securityeventtable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch
            }
        }       
       Section -Name 'UPN Suffixes' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         Panel {
                new-htmlTable -HideFooter -HideButtons -DataTable $DomainTable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch
            }
        }    
    }
    
    New-HTMLPage -Name 'Groups' {
        New-HTMLTab -Name 'Groups' -IconSolid user-alt   {

       Section -Name 'Groups Overivew' -HeaderBackGroundColor Teal -HeaderTextAlignment left {
         Panel {
                new-htmlTable -HideFooter -HideButtons -DataTable $TOPGroupsTable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch
            }
        }          
          
       Section -Name 'Active Directory Groups' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         Panel {
                new-htmlTable -HideFooter -DataTable $Table
            }
        }
        
       Section -Name 'Objects in Default OUs' -HeaderBackGroundColor teal -HeaderTextAlignment left {
            Section -Name 'Domain Administrators' -HeaderBackGroundColor teal  {
                new-htmlTable -HideFooter -DataTable $DomainAdminTable 
                
            }
            Section -Name 'Enterprise Administrators' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $EnterpriseAdminTable
            }
}                  

       New-HTMLSection -HeaderText 'Active Directory Groups Chart' -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Types' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                    $GroupTypetable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Custom vs Default Groups' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    $DefaultGrouptable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Membership' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette3 
                    $GroupMembershipTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Protected From Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4
                    $GroupProtectionTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

        }                
        }
    }

    New-HTMLPage -Name 'OU' {
        New-HTMLTab -Name 'Organizational Units' -IconRegular folder {          
          
       Section -Name 'Organizational Units infos' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         Panel {
                new-htmlTable -HideFooter -DataTable $OUTable
            }
        }
      
                
       New-HTMLSection -HeaderText "Organizational Units Charts" -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'OU Gpos Links' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                    $OUGPOTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'Organizations Units Protected from deletion' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    $OUProtectionTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

        }                

    }

    }

    New-HTMLPage -Name 'GPO' {
        New-HTMLTab -Name 'Group Policy' -IconRegular hourglass {
        
       Section -Name 'Users Overivew"' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         Panel {
                new-htmlTable  -DataTable $GPOTable 
            }
        }

    }


    }

    New-HTMLPage -Name 'Users' {

        New-HTMLTab -Name 'Users' -IconSolid audio-description  {
        
       Section -Name 'Users Overivew' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         Panel {
                new-htmlTable -HideFooter -HideButtons  -DataTable $TOPUserTable -DisableSearch
            }
        }
       
       Section -Name 'Active Directory Users' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         Panel {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -DataTable $UserTable -DefaultSortColumn Name -HideFooter
            }
        }        
        
       Section -Name 'Expiring Items' -HeaderBackGroundColor teal -HeaderTextAlignment left {

            Section -Name "Users Haven't Logged on in $Days Days or more" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                New-HTMLTable -HideFooter -DataTable $userphaventloggedonrecentlytable 
            }
            Section -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                New-HTMLTable -HideFooter -DataTable $NewCreatedUsersTable
            }

        }

       Section -Name 'Accounts' -HeaderBackGroundColor teal -HeaderTextAlignment left {

       Section -Name "Users with Passwords Expiring in less than $DaysUntilPWExpireINT days" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $PasswordExpireSoonTable
            }
       Section -Name "Accounts Expiring Soon" -HeaderBackGroundColor teal -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $ExpiringAccountsTable
            }

        }

       New-HTMLSection -HeaderText "Users Charts" -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Enable Vs Disable Users' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette2
                    $EnabledDisabledUsersTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Password Expiration' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    $PasswordExpirationTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Users Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    $ProtectedUsersTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }
                }
            }

        }
    }


    }

    New-HTMLPage -Name 'Computers' {
        New-HTMLTab -Name 'Computers' -IconBrands microsoft {
        
       Section -Name 'Computers Overivew' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         Panel {
                New-HTMLTable -HideFooter -HideButtons -DataTable $TOPComputersTable
            }
        }
       
         Section -Name 'Computers' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         Panel -Invisible {
                New-HTMLTableOption -DataStore HTML -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-HTMLTable -DataTable $ComputersTable  
                #New-HTMLTab  -DataTable $ComputersTable -DateTimeSortingFormat 'yyyy-MM-dd' -HideFooter 
                            }
            }

          New-HTMLSection -HeaderText 'Computers Charts' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette10 -Mode light
                    $ComputerProtectedTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Enabled Vs Disabled' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4 -Mode light
                    $ComputersEnabledTable.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }

            }

         New-HTMLSection -HeaderText 'Computers Operating System Breakdown' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
                New-HTMLPanel {
                New-HTMLChart -Title 'Computers Operating Systems' -TitleAlignment center  { 
                    New-ChartTheme  -Mode light
                    $GraphComputerOS.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }                    
                }
            }
         
        }


    }


    }

    New-HTMLPage -Name 'Resume'  {
    
    New-HTMLTab -Name 'Resume' {     

       New-HTMLSection -Name 'Graphes' -HeaderBackGroundColor teal -HeaderTextAlignment left  {

            New-HTMLSection -Name 'Nombres d objets' -HeaderBackGroundColor teal {
                new-htmlTable -HideFooter -DataTable $Allobjects
            }
            New-HTMLSection -HeaderText 'All Members' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Pourcent By AD Objects' -TitleAlignment center -Height 300  {
                    New-ChartTheme -Mode light
                    
		    $Allobjects.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }


            }

        }

       Section -Name 'About' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
         New-HTMLPanel {
         New-HTMLList {
              New-HTMLListItem -Text 'Resume All objects AD' 
              New-HTMLListItem -Text "Generated date $time"
              New-HTMLListItem -Text 'Active Directory _ OverHTML  Version : 2.0  Author Dakhama Mehdi - Date : 08/12/2022<br> 
              <br> Inspired ADReportHTLM Version : 1.0.3 Author: Bradley Wyatt - Date: 12/4/2018 [thelazyadministrato](https://www.thelazyadministrator.com/)<br>
              <br> Thanks : JBear,jporgand<br>
              <br> Credit : Mahmoud Hatira, Zouhair sarouti<br>
              <br> Thanks : Boss PrzemyslawKlys - Module PSWriteHTML- [Evotec](https://evotec.xyz) '
              } -FontSize 12
            }
            
          New-HTMLPanel {
            New-HTMLImage -Source $RightLogo 
        } 
        }   
    }

    }    

    
} 
