function HTMLMultiPage {
#region generatehtml

$time = (get-date)

Write-Host "Working on HTML Report ..." -ForegroundColor Green

New-HTML -TitleText 'AD_ModernReport' -ShowHTML -Online -FilePath $SavePath {
   
    New-HTMLNavTop -Logo $CompanyLogo -MenuColorBackground 	gray  -MenuColor Black -HomeColorBackground gray -HomeLinkHome {
       
        New-NavTopMenu -Name 'Domains' -IconRegular address-book -IconColor black  {
        New-NavLink -IconSolid users -Name 'Groups' -InternalPageID 'Groups'
        New-NavLink -IconSolid users -Name 'Groups_Empty' -InternalPageID 'Groups_Empty'    
        New-NavLink -IconMaterial folder -Name 'OU' -InternalPageID 'OU'
        New-NavLink -IconSolid scroll -Name 'Group Policy' -InternalPageID 'GPO'
        New-NavLink -IconSolid print -Name 'Printers' -InternalPageID 'Printers'
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
   
     New-HTMLTabStyle  -BackgroundColorActive Teal   
 
      New-HTMLSection  -Name 'Block infos' -Invisible  {

      New-HTMLPanel -Margin 10 -Width "80%" {

      New-HTMLPanel -BackgroundColor silver  {
      New-HTMLText -TextBlock  {
      New-HTMLText -Text  "Domain : $Forest" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "AD Recycle Bin : $ADRecycleBin" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text "FSMO Roles" -Alignment center -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Infra : $InfrastructureMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Rid : $RIDMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "PDC  : $PDCEmulator" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Naming : $DomainNamingMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -Text  "Schema : $SchemaMaster" -Alignment justify -FontSize 15 -FontWeight bold
      New-HTMLText -LineBreak
      
      }
      }

      }

      New-HTMLPanel -Margin 10  {
      
      New-HTMLPanel -BackgroundColor lightgreen -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $UserDisabled -Alignment justify -FontSize 25 -FontWeight bold 
      New-HTMLText -Text 'Disabled Users' -Alignment justify -FontSize 15 
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-slash fa-3x" } 
      } 
      }

      New-HTMLPanel -BackgroundColor yellowgreen -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock { 
      New-HTMLText -Text $userinactive -Alignment justify -FontSize 25 -FontWeight bold 
      New-HTMLText -Text 'Users not login in Last 90 Days' -Alignment justify -FontSize 15 
      New-HTMLTag -Tag 'span' -Attributes @{ class = "fas fa-user-clock fa-3x" } 
      }
        }
        
         }

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor lightpink  -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $neverlogedenabled -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Users Never logged' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-house-user fa-3x" } 
      }
      }

      New-HTMLPanel -BackgroundColor palevioletred  -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $usercomputerdeleted -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Users/computer in RecycleBin' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-trash-alt fa-3x" } 
      }
      }

}

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor lightblue  -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($DomainAdminTable.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'domain admins' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-edit fa-3x" } 
      }
        }


      New-HTMLPanel -BackgroundColor steelblue -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($EnterpriseAdminTable.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Enterprise Admins' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user-tie fa-3x" } 
      }
      }
      

      }     

      New-HTMLPanel -Margin 10  {

      New-HTMLPanel -BackgroundColor bisque -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $ComputerNotSupported -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Computers Not Supported' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-laptop-medical fa-3x" } 
      }
      }

      New-HTMLPanel -BackgroundColor orange -AlignContentText right -BorderRadius 10px  {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($gponotlinked.count) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'GPOs not Linked' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-scroll fa-3x" } 
      }
      }

}  

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor paleturquoise -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $($ExpiringAccountsTable) -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Expired Account and still Enabled' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-umbrella-beach fa-3x" } 
      }
      }
      
      New-HTMLPanel -BackgroundColor mediumaquamarine -AlignContentText right -BorderRadius 10px {
      New-HTMLText -TextBlock {
      New-HTMLText -Text $expiredsoon -Alignment left -FontSize 25 -FontWeight bold
      New-HTMLText -Text 'Account Expired Soon' -Alignment left -FontSize 15
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-bell fa-3x" } 
      }
      }

      }  

        }
         
      New-HTMLSection  -HeaderBackGroundColor teal -HeaderTextAlignment left  {

      New-HTMLSection -Name 'Created Machines / Users By date in last 30 Days' -Invisible  {
      
      New-HTMLPanel  {
                   New-HTMLChart -Title 'Created Machines / Users By date in last 30 Days' -TitleAlignment center -Height 280 {                 
                    New-ChartAxisX -Names $(($barcreateobject).date)
                    New-ChartLine -Name 'User created' -Value $(($barcreateobject).Nbr_users)
                    New-ChartLine -Name 'PC Created' -Value $(($barcreateobject).Nbr_PC)                  
                }
            }    
         }
      
      New-HTMLSection -HeaderBackGroundColor Teal -Invisible -Width "70%" {    

      New-HTMLPanel  {

                New-HTMLChart -Title 'Created Objects VS Deleted' -TitleAlignment center -Height "100%" {
                    New-ChartToolbar -Download pan
                    New-ChartBarOptions -Gradient -Vertical
                    New-ChartLegend -Name 'Created users', 'Created Machines', 'Deleted Users/machines' 
                    New-ChartBar -Name 'Result Current 30 Days' -Value $lastcreatedusers, $lastcreatedpc, $deletedobject
                }
            }  
      
      New-HTMLSection -Name 'Objects in Default OU'  -Width "80%"  {
            New-HTMLChart -Gradient  {
                New-ChartLegend -LegendPosition bottom 
                New-ChartDonut -Name 'Users' -Value $DefaultUsersinDefaultOUTable.Count
                New-ChartDonut -Name 'Computers' -Value $DefaultComputersinDefaultOU
            }
        }

            }     

        }
      
   New-HTMLSection -Invisible {

      New-HTMLSection -Name "Last Locked Users" -HeaderBackGroundColor DarkGreen  -HeaderTextAlignment left {
                new-htmlTable -HideFooter -DataTable $Unlockusers -HideButtons -DisableSearch
            }

      New-HTMLSection -Name 'UPN Suffix' -HeaderTextAlignment center -HeaderBackGroundColor Black -Width "60%"  {
                New-HTMLTable -DataTable $DomainTable -HideButtons -DisableInfo -DisableSearch -HideFooter -TextWhenNoData 'Information: No UPN Suffixes were found'
      }
    
      New-HTMLSection -Width "60%" -HeaderBackGroundColor Teal -name 'Groups Without members'  {
      
 
      New-HTMLGage -Label 'Empty Groups' -MinValue 0 -MaxValue $totalgroups -Value $Groupswithnomembership -ValueColor Black -LabelColor Black -Pointer
      
            }

      New-HTMLSection -Name "Accounts Created in $UserCreatedDays Days or Less" -HeaderBackGroundColor DarkBlue {
                new-htmlTable -HideFooter -DataTable $NewCreatedUsersTable -DisableInfo -HideButtons -PagingLength 6 -DisableSearch -TextWhenNoData 'Information: No new users have been recently created'
            }



        }


 New-HTMLSection -Name 'Objects in Default OUs' -Invisible  {

      New-HTMLSection -Name 'AD Objects in Recycle Bin' -HeaderBackGroundColor skyblue -Width "70%" {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-htmlTable -HideFooter -DataTable $ADObjectTable -PagingLength 12 -Buttons csvHtml5 
           } 


      New-HTMLSection -Name 'Computers in default OU' -HeaderBackGroundColor teal   {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
                New-htmlTable -HideFooter -DataTable $DefaultComputersinDefaultOUTable -PagingLength 12 -HideButtons 
            }

      New-HTMLSection -Name 'Users in Default OU' -HeaderBackGroundColor brown {
               New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ','
               New-HTMLTable -HideFooter -DataTable $DefaultUsersinDefaultOUTable -PagingLength 12 -HideButtons 
            }

        }    
                   
   
          }

    New-HTMLPage -Name 'Groups' {

        New-HTMLTab -Name 'Groups' -IconSolid user-alt   {

       New-HTMLSection -Invisible {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons -DataTable $TOPGroupsTable -DisablePaging -DisableSelect -DisableStateSave -DisableInfo -DisableSearch 
            }
        }          
          
       New-HTMLSection -Name 'Active Directory Groups With Members' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' -NumberAsString -BoolAsString
                new-htmlTable -HideFooter -DataTable $Table -TextWhenNoData 'Information: No Groups were found'
            }
        }
        
       New-HTMLSection -HeaderText 'Active Directory Groups Chart' -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Types' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                     New-ChartPie -Name 'Security Groups' -Value $SecurityCount
                     New-ChartPie -Name 'Distribution Groups' -Value $DistroCount                                    
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Custom vs Default Groups' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name 'Custom Groups' -Value $CustomGroup
                    New-ChartPie -Name 'Default Groups' -Value $DefaultGroup
                  }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Membership' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette3 
                    New-ChartPie -Name 'With Members' -Value $Groupswithmemebrship
                    New-ChartPie -Name 'No Members' -Value $Groupswithnomembership  
                }
            }

            New-HTMLPanel -Invisible {
                New-HTMLChart -Gradient -Title 'Group Protected From Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4
                    New-ChartPie -Name 'Not Protected' -Value $GroupsNotProtected
                    New-ChartPie -Name 'Protected' -Value $GroupsProtected                   
                }
            }

        } 
               
       }                
     }
    
    New-HTMLPage -Name 'Groups_Empty' {

       New-HTMLTab -Name 'Groups Without Members' -IconSolid print {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $Groupsnomembers
            }
        }


    }
    }    

    New-HTMLPage -Name 'OU' {
     
       New-HTMLTab -Name 'Organizational Units' -IconRegular folder {          
          
       New-HTMLSection -Name 'Organizational Units infos' -Invisible {
         New-HTMLPanel {
                new-htmlTable -HideFooter -DataTable $OUTable -TextWhenNoData 'Information: No OUs were found'
            }
        }      
                
       New-HTMLSection -HeaderText "Organizational Units Charts" -HeaderBackGroundColor teal -HeaderTextAlignment left {
           
            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'OU Gpos Links' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette2 
                    New-ChartPie -Name "OUs with GPO's linked" -Value $OUwithLinked
                    New-ChartPie -Name "OUs with no GPO's linked" -Value $OUNotProtected                                      
                }
            }

            New-HTMLPanel  {
                New-HTMLChart -Gradient -Title 'Organizations Units Protected from deletion' -TitleAlignment center -Height 200  {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name "Protected" -Value $OUProtected
                    New-ChartPie -Name "Not Protected" -Value $OUwithnoLink
                }
            }

        }                

    }

    }

    New-HTMLPage -Name 'GPO' {
        New-HTMLTab -Name 'Group Policy' -IconRegular hourglass {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $GPOs
            }
        }
       
       New-HTMLSection -Invisible {

       New-HTMLSection -name 'Unlinked Details' -HeaderBackGroundColor Teal {
              New-HTMLTable -DataTable $gponotlinked 
       }

       New-HTMLSection -Name 'Linked Vs Unliked GPOs' -HeaderBackGroundColor Teal  {
            New-HTMLChart {
                New-ChartLegend -LegendPosition bottom 
                New-ChartBarOptions -Gradient
                New-ChartDonut -Name 'Unlinked' -Value $gponotlinked.Count -Color silver
                New-ChartDonut -Name 'Linked' -Value $GPOs.Count -Color orange
            }
        }


    }


    }
    }

    New-HTMLPage -Name 'Printers' {

       New-HTMLTab -Name 'Printer server' -IconSolid print {
        
       New-HTMLSection -Name 'Informations"' -Invisible  {
         New-HTMLPanel {
                new-htmlTable  -DataTable $printers
            }
        }


    }
    }

    New-HTMLPage -Name 'Users' {

       New-HTMLTab -Name 'Users' -IconSolid audio-description  {
        
       New-HTMLSection -Name 'Users Overivew' -Invisible  {
         New-HTMLPanel {
                new-htmlTable -HideFooter -HideButtons  -DataTable $TOPUserTable -DisableSearch
            }
        }
       
       New-HTMLSection -Name 'Active Directory Users' -HeaderBackGroundColor Teal -HeaderTextAlignment left  {
         New-HTMLPanel {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' -NumberAsString -BoolAsString
                New-HTMLTable -DataTable $UserTable -DefaultSortColumn Name -HideFooter 
            }
        }                
       
       New-HTMLSection -HeaderText "Users Charts" -HeaderBackGroundColor DarkBlue -HeaderTextAlignment left  {
           
            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Enable Vs Disable Users' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette2
                    New-ChartPie -Name "Enabled" -Value $UserEnabled
                    New-ChartPie -Name "Disabled" -Value $UserDisabled                    
                }
            }

             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Password Expiration' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name "Password Never Expired" -Value $UserPasswordNeverExpires
                    New-ChartPie -Name "Password Expires" -Value $UserPasswordExpires 
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Users Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette1
                    New-ChartPie -Name "Protected" -Value $ProtectedUsers
                    New-ChartPie -Name "Not Protected" -Value $NonProtectedUsers 
                }
            }

        }
    }


    }

    New-HTMLPage -Name 'Computers' {
    New-HTMLTab -Name 'Computers' -IconBrands microsoft {
        
       New-HTMLSection -Name 'Computers Overivew' -Invisible  {
         New-HTMLPanel {
                New-HTMLTable -HideFooter -HideButtons -DataTable $TOPComputersTable
            }
        }
       
         New-HTMLSection -Name 'Computers' -HeaderBackGroundColor teal -HeaderTextAlignment left {
         New-HTMLPanel -Invisible {
                New-HTMLTableOption -DataStore JavaScript -DateTimeFormat 'yyyy-MM-dd' -ArrayJoin -ArrayJoinString ',' -NumberAsString -BoolAsString
                New-HTMLTable -DataTable $ComputersTable  
                            }
            }

          New-HTMLSection -HeaderText 'Computers Charts' -HeaderBackGroundColor DarkBlue -HeaderTextAlignment left  {
     
             New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Protected from Deletion' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette10 -Mode light
                    New-ChartPie -Name 'Protected' -Value $ComputerProtected
                    New-ChartPie -Name 'Not Protected' -Value $ComputersNotProtected                              
                }
            }

            New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Enabled Vs Disabled' -TitleAlignment center -Height 200 {
                    New-ChartTheme -Palette palette4 -Mode light
                    New-ChartPie -Name 'Enabled' -Value $ComputerEnabled
                    New-ChartPie -Name 'Disabled' -Value $ComputerDisabled                  
                }
            }

            }

          New-HTMLSection -Invisible {

         New-HTMLSection -name 'Potential Win10 End of support' {
                       
           New-HTMLChart {
                New-ChartLegend -LegendPosition bottom
                New-ChartDonut -Name 'Endofsupport' -Value $endofsupportwin
                New-ChartDonut -Name 'Windows 10/11 supported' -Value ($allwin1011 - $endofsupportwin)
            }

         }


         New-HTMLSection -HeaderText 'Computers Operating System Distribution' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
           
                New-HTMLPanel {
                New-HTMLChart -Gradient -Title 'Computers Operating Systems' -TitleAlignment center  { 
                    New-ChartTheme  -Mode light
                    $GraphComputerOS.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count 
                    }                    
                }
            }
         
        }
        }


    }
    }         
    
    New-HTMLPage -Name 'Resume'  {
    
    New-HTMLTab -Name 'Resume' {     

    New-HTMLSection -Invisible {

      New-HTMLSection  -HeaderBackGroundColor Teal -Invisible  {

      New-HTMLPanel -Margin 10  {
      
      New-HTMLPanel -BackgroundColor lightgreen -AlignContentText right {
      New-HTMLText -Text $Allobjects[0].count -Alignment left -FontSize 40 -FontWeight bold 
      New-HTMLText -Text $Allobjects[0].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-users fa-3x" } 
      }

      New-HTMLText -LineBreak 

      New-HTMLPanel -BackgroundColor bisque -AlignContentText right {
      New-HTMLText -Text $Allobjects[1].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[1].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-user fa-3x" } 

        }
        
         }

      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor lightblue  -AlignContentText right  {
      New-HTMLText -Text $Allobjects[2].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[2].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-laptop fa-3x" } 
      }

      New-HTMLText -LineBreak 
      
      New-HTMLPanel -BackgroundColor lightpink  -AlignContentText right  {
      New-HTMLText -Text $Allobjects[3].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[3].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-address-card fa-3x" } 
      }

      }    
      
      New-HTMLPanel -Margin 10 {

      New-HTMLPanel -BackgroundColor khaki  -AlignContentText right  {
      New-HTMLText -Text $Allobjects[4].count -Alignment left -FontSize 40 -FontWeight bold
      New-HTMLText -Text $Allobjects[4].name -Alignment left -FontSize 20
      New-HTMLTag -Tag 'i' -Attributes @{ class = "fas fa-print fa-3x" } 
      }

      New-HTMLText -LineBreak 
      
      }   
        }
      
      New-HTMLSection -HeaderText 'All Members' -Invisible {
     
             New-HTMLPanel -Width "70%" {
                New-HTMLChart -Gradient -Title 'Pourcent By AD Objects' -TitleAlignment center -Height 300  {
                    New-ChartTheme  -Mode light                    
		    $Allobjects.GetEnumerator() | ForEach-Object {
                    New-ChartPie -Name $_.name -Value $_.count
                    }                    
                }
            }


            }

    }       

    New-HTMLSection -Name 'About' -HeaderBackGroundColor teal -HeaderTextAlignment left  {
   
    New-HTMLPanel {
    New-HTMLList {
              New-HTMLListItem -Text 'Modern Active Directory _ Version : 1.0 _ Release : 01/2023' 
              New-HTMLListItem -Text "Generated date : $time"
              New-HTMLListItem -Text 'Author : Dakhama Mehdi<br> 
              <br> Inspired ADReportHTLM Bradley Wyatt - Release 12/4/2018 [thelazyadministrato](https://www.thelazyadministrator.com/)<br>
              <br> Credit : Thirrey Demon-Barcelo, Mattieu Souin, Mahmoud Hatira, Zouhair sarouti<br>
              <br> Thanks : Boss Przemyslaw Klys - Module PSWriteHTML- [Evotec](https://evotec.xyz)'
              } -FontSize 14
            }         
    New-HTMLPanel {
            New-HTMLImage -Source $RightLogo 
        } 
        }   
    }

    }     

    
} 

#endregion generatehtml
}