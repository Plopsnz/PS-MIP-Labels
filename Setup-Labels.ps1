<#
Author: simon.rowan@theinstillery.com
Date: 25/03/2022
Tag: MIP sensitivity Labels
Notes: This will get you going, once you have run through the commands below you can go into the Compliance portal and start tweaking the labels and the policies that apply them to suite your requirements.
Disclaimer: Please refer to the Readme.md file
#>



#Variables
$UserCredential = Get-Credential #Get Credentials for Azure connection to login as a admin


#Connect to Exchange
#if you don't have it installed you will need to run "Install-Module ExchangeOnlineManagement" first
Import-Module ExchangeOnlineManagement


#Create the session
Connect-IPPSSession -Credential $UserCredential

<# Or user this format if you have trouble with MFA
Connect-ExchangeOnline -UserPrincipalName user@domain.onmicrosoft.com
Connect-IPPSSession -UserPrincipalName user@domain.onmicrosoft.com
#>

#create personal label
New-Label -DisplayName "Personal" -Name "Personal" -Tooltip "Use This classification to identify information that is not business related" -Comment "Use This classification to identify information that is not business related"

#create public label
New-Label -DisplayName "Public" -Name "Public" -Tooltip "Use This classification to identify information that is business related and is for public consumption" -Comment "Use This classification to identify information that is business related and is for public consumption" -ContentType "Site, UnifiedGroup"

#create General label
New-Label -DisplayName "General" -Name "General" -Tooltip "Use This classification to identify information that is business related and is ok for general business use" -Comment "Use This classification to identify information that is business related and is ok for general business use" -ContentType "Site, UnifiedGroup"

#create public label
New-Label -DisplayName "Confidential" -Name "Confidential" -Tooltip "Use This classification to identify information that is business related and needs to be protected" -Comment "Use This classification to identify information that is business related and needs to be protected" -ContentType "Site, UnifiedGroup"

<#
I have broken out the label creation above so it is easy to read and understand what you are doing, you could how ever do the same thing many differnet ways.
See below another way to do the same thing, it looks better, would run a little faster (maybe), so its up to you :)

$labelnamedata = '{
        "LabelName":"Personal",
        "LabelDescription":"Use This classification to identify information that is not business related",
        "LabelTooltip":"Use This classification to identify information that is not business related"
    },  
    {
        "LabelName":"Public",
        "LabelDescription":"Use This classification to identify information that is business related and is for public consumption",
        "LabelTooltip":"Use This classification to identify information that is business related and is for public consumption"
    },  
    {
        "LabelName":"General",
        "LabelDescription":"Use This classification to identify information that is business related and is ok for general business use",
        "LabelTooltip":"Use This classification to identify information that is business related and is ok for general business use"
    },  
    {
        "LabelName":"Confidential",
        "LabelDescription":"Use This classification to identify information that is business related and needs to be protected",
        "LabelTooltip":"Use This classification to identify information that is business related and needs to be protected"
    }'
 
$labels = $labelnamedata | ConvertFrom-Json
 
foreach($label in $labels)
{
    $createlabel = New-Label `
                            -DisplayName $label.LabelName `
                            -Name $label.LabelName `
                            -Comment $label.LabelDescription `
                            -Tooltip $label.LabelTooltip
}

#>



#add sublabels to parents
New-Label -DisplayName "Internal" -Name "Internal" -Tooltip "Use to identify and protect Confidential data that is be shared inside the Organization" -Comment "Use to identify and protect Confidential data that is be shared inside the Organization" -ContentType "File, Email" -ParentId "Confidential"
New-Label -DisplayName "External" -Name "External" -Tooltip "Use to identify and protect Confidential data that is be shared inside the Organization" -Comment "Use to identify and protect Confidential data that is be shared inside the Organization" -ContentType "File, Email" -ParentId "Confidential"

#add footers to internal protected documents
Set-Label -Identity "Internal" -ApplyContentMarkingFooterAlignment "Left" -ApplyContentMarkingFooterEnabled $true -ApplyContentMarkingFooterFontColor "#000000" -ApplyContentMarkingFooterFontName "Courier New" -ApplyContentMarkingFooterFontSize "10" -ApplyContentMarkingFooterMargin "5" -ApplyContentMarkingFooterText "Confidential (Internal)"

#add footers to Confidential documents
Set-Label -Identity "External" -ApplyContentMarkingFooterAlignment "Left" -ApplyContentMarkingFooterEnabled $true -ApplyContentMarkingFooterFontColor "#000000" -ApplyContentMarkingFooterFontName "Courier New" -ApplyContentMarkingFooterFontSize "10" -ApplyContentMarkingFooterMargin "5" -ApplyContentMarkingFooterText "Confidential (External)" -EncryptionEnabled $true


#Create the label policay that will publish the labels to the users

New-LabelPolicy -Name "Global" -Labels "Personal","Public","General","Confidential","Internal","External"


