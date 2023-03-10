#This example limit groups searcher to 400 instead 200

import-module ModernActiveDirectory

Get-ADModernActiveDirectory -maxsearchergroups 400