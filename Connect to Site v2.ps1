
############################################
# SharePoint Unattended Build 
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

#Enable Server Publishing guid 94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb
$publishingSP = '94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb'

#Enable Wiki guid 00bfea71-d8fe-4fec-8dad-01c19a6e4053
$wiki = '00bfea71-d8fe-4fec-8dad-01c19a6e4053'


    
#Title for Menu        
$title = "SharePoint Script"
$quit = "Press 9 to Crtl+C to exit"
$menu = "                               
`nSelect the following options
`nPress 1 - Verfiy SharePoint Pnp Module
`nPress 2 - Connect to SharePoint Page
`nPress 3 - Enable Site Features (Collaboration, Enterprise, Feed, Pages, Publishing, & Wiki )
`nPress 4 - Get current Site Features
`nPress 5 - Create Example Site
`nPress 6 - Get all List pages
`nPress 7 - Convert all List pages to Modern Experinces
`nPress 8 - Disconnect for SharePoint Page


"

       
        
#Looping statement until var $selection ends
do {
        
    $Path = $PSScriptRoot
    Write-Host "$title,`n$quit" -ForegroundColor Yellow 
    $selection = Read-Host -Prompt "$menu"
   
                
    # Iterating through loop
    Try { 
        switch ([int]$selection) {
                        
            '1' { Get-Module -Name SharePointPnPPowerShellOnline } 

            '2' {
                Clear-Host
                $site = Read-Host -Prompt " `n Enter SharePoint Site URL (copy and paste)" ; 
                Connect-PnPOnline -Url $site -UseWebLogin ;
                Write-Host " `n Connected To" ; 
                Get-PnPSite ;
            }

            '3' { 

                Clear-Host
                Enable-PnPFeature -Identity $publishing ;
                Enable-PnPFeature -Identity $standard ;
                Enable-PnPFeature -Identity $enterprise ;
                Enable-PnPFeature -Identity $feed ;
                Enable-PnPFeature -Identity $pages ;
                Enable-PnPFeature -Identity $collab ;
                Enable-PnPFeature -Identity $publishingSP ;
                Enable-PnPFeature -Identity $wiki ;
            }

            '4' {
                Clear-Host
                Write-Host "Current Site Feature" ;  
                Get-PnPFeature ;
            }

            '5' {
                #var for example comments
                $entry1 = "This is a text web part in a one column section. You can click inside this text block when in Edit mode to make changes. Besides being able to change the text, you can use the toolbar to choose styles, make other format changes, and add links. You can even add a table by clicking the ellipses (…) in the toolbar for more options."
                $entry2 = " Hello! This is a Text web part in one of two columns in this section. You can click inside this text block when in Edit mode to make changes. Next to this paragraph is a column that contains an image web part. Click the image, and you can use the toolbar to change the image, add a link, crop the image, and more. Learn more about the text web part and the image web part. When you're done editing this page, you can click Save as draft to save your changes and leave edit mode. Only people with edit permissions on your site will be able to see it. If you are ready to make this page visible to everyone who can view your site, click Publish or Post news. For more information, see What happens when I publish a page?  " 
                $entry3 = "This is the other web part of the two column"


                #add client side modern page
                Add-PnPClientSidePage -Name "Example-Page" ;

                #adding section within page
                Add-PnPClientSidePageSection -Page "Example-Page" -SectionTemplate OneColumn  ; 
                Add-PnPClientSidePageSection -Page "Example-Page" -SectionTemplate TwoColumn  ; 

                #inserting texts
                Add-PnPClientSideText -Page "Example-Page" -Text $entry1 -Section 1 -Column 1 ;
                Add-PnPClientSideText -Page "Example-Page" -Text $entry2 -Section 2 -Column 1 ;
                Add-PnPClientSideText -Page "Example-Page" -Text $entry3 -Section 2 -Column 2 ;
            }

            '6' { 
                Write-Host "Current List of SharePoint" ;
                Get-PnPList | select Title, ListExperienceOptions | Out-GridView
            }   

            '7' {

                $SharePointList = Get-PnPList | foreach { $_.Id }
                foreach ($list in $SharePointList) {
                    Set-PnPList -Identity $list -ListExperience NewExperience -ErrorAction SilentlyContinue
                }

            }

            '8' { 
                
                Write-Host " `nDisconnecting From SharePoint Session" -ForegroundColor Red
                Disconnect-PnPOnline 
            
            }
        
        }

    }

    catch {            
        Clear-Host
        Wait-Event -Timeout 0.5
        Write-Host "Restart Powershell script" -ForegroundColor Red; 
    }


} until ($selection -eq 9)



                
