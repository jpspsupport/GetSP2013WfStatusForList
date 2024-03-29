﻿<#
 This Sample Code is provided for the purpose of illustration only and is not intended to be used in a production environment. 
 THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
 INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  

 We grant you a nonexclusive, royalty-free right to use and modify the sample code and to reproduce and distribute the object 
 code form of the Sample Code, provided that you agree: 
    (i)   to not use our name, logo, or trademarks to market your software product in which the sample code is embedded; 
    (ii)  to include a valid copyright notice on your software product in which the sample code is embedded; and 
    (iii) to indemnify, hold harmless, and defend us and our suppliers from and against any claims or lawsuits, including 
          attorneys' fees, that arise or result from the use or distribution of the sample code.

Please note: None of the conditions outlined in the disclaimer above will supercede the terms and conditions contained within 
             the Premier Customer Services Description.
#>
param(
  [Parameter(Mandatory=$true)]
  $siteUrl,
  [Parameter(Mandatory=$true)]
  $listName,
  $username,
  $password,
  $outfile,
  $resume = $false
)

Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "C:\Program Files\Common Files\microsoft shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.WorkflowServices.dll"

$script:context = new-object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$pwd = $null
if ($username -eq $null)
{
  $cred = Get-Credential
  $username = $cred.UserName
  $pwd = $cred.Password
}
else
{
  $pwd = convertto-securestring $password -AsPlainText -Force
}
$credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $pwd)
$script:context.Credentials = $credentials


$script:context.add_ExecutingWebRequest({
  param ($source, $eventArgs);
  $request = $eventArgs.WebRequestExecutor.WebRequest;
  $request.UserAgent = "NONISV|Contoso|Application/1.0";
})

function ExecuteQueryWithIncrementalRetry {
  param (
      [parameter(Mandatory = $true)]
      [int]$retryCount
  );

  $DefaultRetryAfterInMs = 120000;
  $RetryAfterHeaderName = "Retry-After";
  $retryAttempts = 0;

  if ($retryCount -le 0) {
      throw "Provide a retry count greater than zero."
  }

  while ($retryAttempts -lt $retryCount) {
      try {
          $script:context.ExecuteQuery();
          return;
      }
      catch [System.Net.WebException] {
          $response = $_.Exception.Response

          if (($null -ne $response) -and (($response.StatusCode -eq 429) -or ($response.StatusCode -eq 503))) {
              $retryAfterHeader = $response.GetResponseHeader($RetryAfterHeaderName);
              $retryAfterInMs = $DefaultRetryAfterInMs;

              if (-not [string]::IsNullOrEmpty($retryAfterHeader)) {
                  if (-not [int]::TryParse($retryAfterHeader, [ref]$retryAfterInMs)) {
                      $retryAfterInMs = $DefaultRetryAfterInMs;
                  }
                  else {
                      $retryAfterInMs *= 1000;
                  }
              }

              Write-Output ("CSOM request exceeded usage limits. Sleeping for {0} seconds before retrying." -F ($retryAfterInMs / 1000))
              #Add delay.
              Start-Sleep -m $retryAfterInMs
              #Add to retry count.
              $retryAttempts++;
          }
          else {
              throw;
          }
      }
  }

  throw "Maximum retry attempts {0}, have been attempted." -F $retryCount;
}


function EnumWorkflowsInFolder($list, $ServerRelativeUrl)
{
  do
  {
    $camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
    $camlQuery.ListItemCollectionPosition = $position
    $camlQuery.ViewXml = "<View><RowLimit>5000</RowLimit></View>";
    if ($serverRelativeUrl -ne $null)
    {
       $camlQuery.FolderServerRelativeUrl = $ServerRelativeUrl
    }
    $listItems = $list.GetItems($camlQuery);
    $script:context.Load($listItems);
    ExecuteQueryWithIncrementalRetry -retryCount 5 

    foreach($listItem in $listItems)
    {
      if ($listItem.FileSystemObjectType -eq [Microsoft.SharePoint.Client.FileSystemObjectType]::Folder)
      {
         $script:context.Load($listItem.Folder)
         ExecuteQueryWithIncrementalRetry -retryCount 5 
         EnumWorkflowsInFolder -List $list -ServerRelativeUrl $listItem.Folder.ServerRelativeUrl
      }

      $wfic = $script:wfis.EnumerateInstancesForListItem($list.Id, $listItem.Id);
      $script:context.Load($wfic);
      ExecuteQueryWithIncrementalRetry -retryCount 5 
 
      foreach ($wfi in $wfic)
      {
        WriteOut -text (($listItem.Id.ToString()) + "," + (GetWorkflowSubscription -subid $wfi.WorkflowSubscriptionId) + "," + $wfi.Status) -append $true
        if ($resume)
        {
            if ($wfi.Status -eq "Suspended")
            {
                $script:wfis.ResumeWorkflow($wfi)
                ExecuteQueryWithIncrementalRetry -retryCount 5 
                Write-Output ("Resumed workflow on Item ID = " + $listItem.Id.ToString())
            }
        }
      }
    } 
    $position = $listItems.ListItemCollectionPosition
  }
  while($position -ne $null)
}

function GetWorkflowSubscription($subid)
{
  if ($script:wfsubhash[$subid.ToString()] -eq $null)
  {    
    $sub = $script:wfss.GetSubscription($subid)
    $script:context.Load($sub)
    ExecuteQueryWithIncrementalRetry -retryCount 5 

    $script:wfsubhash[$subid.ToString()] = $sub.Name
    return $sub.Name
  }
  else
  {
    return $script:wfsubhash[$subid.ToString()]
  }
}

function WriteOut($text, $append)
{
  if ($outfile -eq $null)
  {
    Write-Output $text
  }
  else
  {
    if ($append)
    {
      $text | Out-File $outfile -Append -Encoding UTF8
    }
    else
    {
      $text | Out-File $outfile -Encoding UTF8
    }
  }
}


$wfsm = new-object Microsoft.SharePoint.Client.WorkflowServices.WorkflowServicesManager($script:context, $script:context.Web)
$script:wfss = $wfsm.GetWorkflowSubscriptionService();
$script:wfsubhash = @{}
$script:wfis = $wfsm.GetWorkflowInstanceService();

$list = $script:context.Web.Lists.GetByTitle($listName)
$script:context.Load($list)
ExecuteQueryWithIncrementalRetry -retryCount 5 

$wfSubs = $wfss.EnumerateSubscriptionsByList($list.Id);
$script:context.Load($wfSubs);
ExecuteQueryWithIncrementalRetry -retryCount 5 

WriteOut -text "ItemId,WorkflowName,Status" 
EnumWorkflowsInFolder -List $list -serverRelativeUrl $null