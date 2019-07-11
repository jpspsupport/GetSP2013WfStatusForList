# GetSP2013WfStatusForList 

This is a PowerShell implementation of enumerating the SharePoint 2013 mode workflow for a specified list.

This code has a prerequisite of ShraePoint Online CSOM.

## Prerequitesite
You need to download SharePoint Online Client Components SDK to run this script.
https://www.microsoft.com/en-us/download/details.aspx?id=42038

You can also acquire the latest SharePoint Online Client SDK by Nuget as well.

1. You need to access the following site. 
https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM

2. Download the nupkg.
3. Change the file extension to *.zip.
4. Unzip and extract those file.
5. place them in the specified directory from the code. 

## How to Run - parameters

-siteUrl ... Target site collection (site) or site (web) URL.
-listName ... Target List Name (Title)
-username ... Site Administrator Account to check the workflow instances.
-password ... The password of the above user.
-output   ... [optional] When parameter is specified console ouput is saved to file instead.
-resume   ... [optional] When resume is set to $true, the suspended workflows are resumed when detected.

#### Example 1: Displaying all the workflow instances and their states
.\GetSP2013WfStatusForList.ps1 -siteUrl https://tenant.sharepoint.com/sites/wf -listName customlist -username admin@tenant.onmicrosoft.com -password PASSWORD

#### Example 2: Writing all the workflow instances and their states to files
.\GetSP2013WfStatusForList.ps1 -siteUrl https://tenant.sharepoint.com/sites/wf -listName customlist -username admin@tenant.onmicrosoft.com -password PASSWORD -outfile .\output.csv

#### Example 3: Resuming all the suspended instances
.\GetSP2013WfStatusForList.ps1 -siteUrl https://tenant.sharepoint.com/sites/wf -listName customlist -username admin@tenant.onmicrosoft.com -password PASSWORD -resume $true








## Reference
Original Documentation of this sample code from the following Japan SharePoint Support Team's forum.
https://blogs.technet.microsoft.com/sharepoint_support/2018/08/10/sample-sharepoint-2013-workflow-status/
