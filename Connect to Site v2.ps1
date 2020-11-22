
############################################
# SharePoint Unattended Site Build 
# Author: ahmad.buhari@us.af.mil
# Date: November 2020
############################################






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
        `n Press 2 - Enter Site Name
        `n Press 3 - Enable Site (SharePoint Server Enterprise)
        `n Press 4 - 
        `n Press 5 - Quit   
     
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
                        
                        '1' {Get-Module -Name SharePointPnPPowerShellOnline} 
                        '2' {$site = Read-Host -Prompt 'Enter SharePoint Site' (Connect-PnPOnline -Url $site -UseWebLogin)}
                        '3' {Connect-PnPOnline -Url $site -UseWebLogin}
                        '5' {Write-Host " `n Exiting Window " -ForegroundColor Red}
        
                        }

                    }

                        catch { [System.OutOfMemoryException]
                    
                        Write-Host 'Restart Powershell script' -ForegroundColor Red
                                }
                
                            if ($selection -gt 5) {Write-Host 'try again'}

                
                            } until ($selection -eq 5)

                
            

#Site Feature Identity

#Publishing guid 22a9ef51-737b-4ff2-9346-694633fe4416

    #$Publishing = '22a9ef51-737b-4ff2-9346-694633fe4416'

    #Write-Host Enabling Publising -ForegroundColor Yellow 

    #Enable-PnPFeature -Identity $Publishing 



#Enable SharePoint Standard Site feature 99fe402e-89a0-45aa-9163-85342e865dc8 (Display Name: BaseWeb)

#Enable Enterprise SharePoint Site feature 0806d127-06e6-447a-980e-2e90b03101b8 (Premium Web)

#Enable Site Feeds guid 15a572c6-e545-4d32-897a-bab6f5846e18

#Enable site pages guid b6917cb1-93a0-4b97-a84d-7cf49975d4ec

#Enable SharePoint Publishing guid dffaae84-60ee-413a-9600-1cf431cf0560 (RollupPages)

#Enable Publishing web guid 94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb

#Enable WebFeature a0e5a010-1329-49d4-9e09-f280cdbed37d

#Enable Site Feeds guid 15a572c6-e545-4d32-897a-bab6f5846e18

#Enable SharePoint Publishing guid dffaae84-60ee-413a-9600-1cf431cf0560 (RollupPages)

#Enable Team Collab guid  00bfea71-4ea5-48d4-a4ad-7ea5c011abe5