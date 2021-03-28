####Feature10 = Project Document Folder in-progress####

#Naming new sharepoint list
#Lessons: Spaces are not friendly when redirecting to a web browser
$site = Read-Host -Prompt "Enter SharePoint Site"
$project = Read-Host -Prompt "Enter Project Name"
$project0 = "$project-List"
$new = New-PnPList -Title "$project0" -Template GenericList 

Write-Host "Building List..." -ForegroundColor Yellow



#Lessons: Identical xml attribute can only be assigned at a single time. Using variables to separate each new column

#Update column with Document (Text) required, Attachment, Link (URL), Category (Text), Type (Text), Submitter (user), Date 

    $col1TaskXml = '<Field Type="Text" Name="Task" DisplayName="Task" Viewable="TRUE" />'


    $col2StatusXml = '
                    <Field Type="Choice" Name="Status" DisplayName="Status" Viewable="TRUE">
                        <CHOICES>
                            <CHOICE>Not Started</CHOICE>
                            <CHOICE>In Progress</CHOICE>
                            <CHOICE>Completed</CHOICE>
                            <CHOICE>Deferred</CHOICE>
                        </CHOICES>
                        <Default>Not Started</Default>
                    </Field>
                '   

#Customizing Column 1
$col1 = Add-PnPFieldFromXml -List "$project0" -FieldXml $TaskXml;

Write-Host "Customizing..." -ForegroundColor Yellow;

#Customizing Column 2
$col2 = Add-PnPFieldFromXml -List "$project0" -FieldXml $StatusXml;

#Clarity: "List" points to the acutal list name within SharePint, "Identity" refers to the "View" name of the list (deafult is "All items"), Fields refers to the columns (multiple coulmns can be called)

#Add Custom View
$view = Set-PnPView -List "$project0" -Identity "All Items" -Fields "Task", "Status" -ErrorAction Ignore ;

Write-Host "Updating New Look..." -ForegroundColor Yellow

#Pauses for code to inject in SharePoint
Start-Sleep -s 2 ;

#Retriving new list url
$project1 = Get-PnPList "$project0" ;

#URL
$full = ($site + $project1.DefaultViewUrl) ;

#opens in web browser
Write-Host "`nNew Project Page Url:" $full -ForegroundColor Yellow 
Write-Host "Loading Page..." -ForegroundColor Yellow

$newList = Start-Process -FilePath Chrome $full
                
