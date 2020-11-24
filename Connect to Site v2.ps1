
############################################
# SharePoint Unattended Site Build 
# Author: ahmad.buhari@us.af.mil
# Date: November 2020
############################################



#Site Feature Identity

#Enable SharePoint Publishing guid dffaae84-60ee-413a-9600-1cf431cf0560 (RollupPages)

$publishing = 'dffaae84-60ee-413a-9600-1cf431cf0560'

#Enable SharePoint Standard Site feature 99fe402e-89a0-45aa-9163-85342e865dc8 (Display Name: BaseWeb)

$standard = '99fe402e-89a0-45aa-9163-85342e865dc8'   

#Enable Enterprise SharePoint Site feature 0806d127-06e6-447a-980e-2e90b03101b8 (Premium Web)

$enterprise = '0806d127-06e6-447a-980e-2e90b03101b8'

#Enable Site Feeds guid 15a572c6-e545-4d32-897a-bab6f5846e18

$feed = '15a572c6-e545-4d32-897a-bab6f5846e18'

#Enable site Pages guid b6917cb1-93a0-4b97-a84d-7cf49975d4ec

$pages = 'b6917cb1-93a0-4b97-a84d-7cf49975d4ec'

#Enable Team Collab guid 00bfea71-4ea5-48d4-a4ad-7ea5c011abe5

$collab = '00bfea71-4ea5-48d4-a4ad-7ea5c011abe5'

#Enable WebFeature a0e5a010-1329-49d4-9e09-f280cdbed37d






##Function for Menu

function PowerShell-Menu {


    #Menu options
    param ([string]$Title = '')

    #clear screen
    Clear-Host

    #Outpput Menu Title & Menu Selection
    Write-Host "---$Title---

        `n
        `n Press 1 - Verfiy SharePoint Pnp Module
        `n Press 2 - Connect to SharePoint Page
        `n Press 3 - Enable Site Feature
        `n Press 4 - Get current Site Features
        `n Press 5 - Example Project Example Site
        `n Press 6 - Create Wiki Example Page
        `n Press 9 - Quit   
     
        " -ForegroundColor Yellow
              
}

    
#Title for Menu        
PowerShell-Menu -Title 'SharePoint Unattended Commands'

       
        
#Looping statement until var $selection ends
do {
        
    $Path = $PSScriptRoot 
    $selection = Read-Host -Prompt "`n Please Select the following options"
        
                
    # Iterating through loop
    Try { 
        switch ($selection) 
        {
                        
            '1' { Get-Module -Name SharePointPnPPowerShellOnline } 

            '2' {
                $site = Read-Host -Prompt " `n Enter SharePoint Site URL (copy and paste)" ; 
                Connect-PnPOnline -Url $site -UseWebLogin ;
                Write-Host " `n Connected To" ; 
                Get-PnPSite 
            }

            '3' { 

                Enable-PnPFeature -Identity $publishing ;
                Enable-PnPFeature -Identity $standard ;
                Enable-PnPFeature -Identity $enterprise ;
                Enable-PnPFeature -Identity $feed ;
                Enable-PnPFeature -Identity $pages ;
                Enable-PnPFeature -Identity $collab ;
            }

            '4' { Write-Host "Current Site Feature" ;  

                Get-PnPFeature ;
            }

            '5' {


            }

            '9' { Write-Host " `n Exiting Window `n " -ForegroundColor Red }
        
        }

    }

    catch {
        [System.OutOfMemoryException]
                    
        Write-Host 'Restart Powershell script' -ForegroundColor Red
    }
                
    if ($selection -gt 9) { Write-Host "Please Choose Again" -ForegroundColor Red }

                
} until ($selection -eq 9)

                