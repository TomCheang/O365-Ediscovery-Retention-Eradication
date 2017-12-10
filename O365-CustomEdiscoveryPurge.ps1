<#
Dec 10, 2016
Custom functions for Ediscovery search for Exchange Online
Limitations of O365 Compliance Search center:
Start-ComplianceSearch -Preview only returns 1000 results.  Generating a full
  report requires manually download that takes several hours.

KQL (Exchange 2013+) cannot target a domain and returns indexed results for
  recipient/sender fields.  Will need to do additional filtering on result set
  returned.

Because KQL returns more broader set of matches due to indexed sender fields,
  the limitation of 1000 items makes built-in Compliance Search unusable for
  intended purpose of targeting specific items by sender.
#>

#Generates kql string for each run
function New-CustomKQL {
  Param (
    [string[]] $Senders,
    [string[]] $ExclusionList,
    [int]$RetentionAge
  )
  [string]$SenderString = ''
  foreach ($s in $Senders) {
    $SenderString += 'From:"' + $s + '" OR '
  }
  [string]$ExclusionString = ''
  foreach ($e in $ExclusionList) {
    $ExclusionString += ' NOT from:"' + $e + '"'
  }
  $sender = '(' + $SenderString.TrimEnd(' OR ') +
    $ExclusionString + ') '

  if ($RetentionAge -gt 0) {
    $RetentionAge = $RetentionAge * -1
  }
  # starting from founding date
  $RetentionDate = (Get-Date).AddDays($RetentionAge).ToShortDateString()
  $DateString = 'AND Sent:2/4/2004..' + $RetentionDate
  [string]$string = ($sender + $DateString + ' AND Kind:email')

  Return $string
}

#Primary function.  Gets full list of searchable mailboxes.  Gets full result
#  set by returning PreviewItems and MailboxStats of  specified batch size.
#  'eDiscoveryMailboxes' is DDL containing all UserMailboxes in Exch Online
function Start-FullCustomEdiscoverySearch {
  [CmdletBinding()]
  Param ($service, $KQL, $MailboxGroupFilter = 'eDiscoveryMailboxes',
    #Prevent max allowable searchable mailboxes
    [ValidateScript({$_.Count -lt '10000'})]
    [int]$BatchSize = '2000',
    [int]$PageSize = '1000'
  )
  $splMailboxGroupFilter = @{
    service = $service
    MailboxGroupFilter = $MailboxGroupFilter
    ExpandMembership = $true
  }
  $FullSearchableSet = Get-AllSearchableMailboxes @splMailboxGroupFilter

  $splBatchSplit = @{
    FullSearchableSet = $FullSearchableSet
    BatchSize = $BatchSize
  }
  $Batches = Split-MailboxSearchScope @splBatchSplit
  $start = Get-Date
  Write-Host "`r`n`r`nBegin Full custom ediscovery process`r`n"
  write-Host "KQL: `r`n$kql"
  Write-Host "`r`nMailboxes: $($FullSearchableSet.count)          " -NoNewline
  Write-Host "Batches: $($Batches.count)         " -NoNewline
  Write-Host "Batch Size: $BatchSize          "
  Write-Host "Start Time:$($start.ToLocalTime()) `r`n`r`n"

  $PreviewItems = New-Object System.Collections.ArrayList
  $MailboxStats = New-Object System.Collections.ArrayList
  $i = 0
  Try {
    foreach ($b in $Batches) {
      $start = Get-Date
      $i++; Write-Host "Batch #: $i         " -NoNewline
      Write-Host "Start Time:$($start.ToLocalTime()) "

      $splBatchedSearch = @{
        service = $service
        KQL = $kql
        Mailboxes = $b
        PageSize = $PageSize
      }
      $BatchResult = Start-BatchCustomEdiscoverySearch @splBatchedSearch
      if ($BatchResult.PreviewItems -or $BatchResult.MailboxStats) {
        $PreviewItems.AddRange($BatchResult.PreviewItems)
        $MailboxStats.AddRange($BatchResult.MailboxStats)
      }
      $end = Get-Date
      Write-Host "Items in Batch: $($BatchResult.PreviewItems.count)        " -NoNewline
      Write-Host "Total PreviewItems: $($PreviewItems.count)"
      Write-Host "End Time:$($end.ToLocalTime())          "
      $Duration = ($end - $start)
      Write-Host "Duration:  $($Duration.Minutes) Min   $($Duration.Seconds) sec"
      Write-Host "`r`n`r`n"
    }
    $TotalRunTime = ((Get-Date) - $start)
    Write-Host "Total Runtime:  $($TotalRunTime.ToString())"
  }
  Catch {
    Write-Host $Error[0]
  }
  Finally {
    $hash = @{
      PreviewItems = $PreviewItems
      MailboxStats = $MailboxStats
    }
    New-Object psobject -Property $hash
  }
}

#Given DL, expands membership to obtain full list of searchable mailboxes
function Get-AllSearchableMailboxes {
  Param (
    $service,
    # DDL 'eDiscoveryMailboxes' contains all cloud mailboxes
    # Group is also in exclusive scope to prevent needless exposure
    $MailboxGroupFilter,
     [ValidateSet($true, $false)]
    $ExpandMembership = $true
  )
  # $true for Expand group members
  $response = $service.GetSearchableMailboxes(
    $MailboxGroupFilter, $ExpandMembership
  )
  if ($response.Result -eq 'Success') {
    Return $response.SearchableMailboxes
  }
  else {
    Return $response.ErrorMessage
  }
}

#Max batch size is 20,000 mbx for Exchange Online.  This function creates an
#  array of mailboxes according to specified batch size
function Split-MailboxSearchScope {
  Param ($FullSearchableSet, [int]$BatchSize)
  #round up number of batches
  $TotalBatches = [System.Math]::Ceiling($FullSearchableSet.count / $BatchSize)
  $SplitJob = New-Object System.Collections.ArrayList
  $i = 0
  for ($b = 0; $b -lt $TotalBatches; $b++) {
    $batch = $FullSearchableSet[$i..($i + $BatchSize -1 )]
    $SplitJob.Add($batch) | Out-Null
    $i += ($BatchSize - 1)
  }
  Return $SplitJob
}

#Returns one paged results given PageSize and PageReferenceItem
function Get-EdiscoveryPagedSearchResults {
  #Returns SearchMailboxesResponse
  Param (
     $service,
    [string]$KQL,
    [ValidateScript({$_.Count -lt '10000'})]
    $Mailboxes,
    [ValidateRange(1,1000)]
    [int] $PageSize = '500',
    [string]$PageItemReference
  )
  $GlobalSearchScope = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope[] $Mailboxes.Length
  $i = 0
  $SearchLocationAll = [Microsoft.Exchange.WebServices.Data.MailboxSearchLocation]::All
  foreach ($m in $Mailboxes) {
    #eDiscovery search takes ReferenceID and not smtpaddress as input
    $mb = New-Object Microsoft.Exchange.WebServices.Data.MailboxSearchScope(
      $m.ReferenceId, $SearchLocationAll
    )
    $GlobalSearchScope[$i] = $mb
    $i++
  }

  $SearchParams = New-Object Microsoft.Exchange.WebServices.Data.SearchMailboxesParameters
  $MbQuery = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery($KQL, $GlobalSearchScope)

  $QueryArray = New-Object Microsoft.Exchange.WebServices.Data.MailboxQuery[] 1
  $QueryArray[0] = $MbQuery
  $SearchParams.SearchQueries = $QueryArray
  $SearchParams.PageSize = $PageSize
  $SearchParams.PageDirection = [Microsoft.Exchange.WebServices.Data.SearchPageDirection]::Next
  $SearchParams.PerformDeduplication = $false
  $SearchParams.ResultType = [Microsoft.Exchange.WebServices.Data.SearchResultType]::PreviewOnly
  if ($PageItemReference) {
    $SearchParams.PageItemReference = $PageItemReference
  }
  $retries = 0
  Do {
    Try {
      $response = $service.SearchMailboxes($SearchParams)
      if ($response.Result -eq 'Success' -and
          $response.SearchResult.PreviewItems.count -gt 0) {
        Return $response.SearchResult
      }
        elseif ($response.Result -eq 'Success' -and
          $response.SearchResult.PreviewItems.count -eq 0) {
        Write-Host "No items found for current batch of mailboxes."
        Return
      }
    }
    Catch [Microsoft.Exchange.WebServices.Data.ServerBusyException] {
      $retries++
      if ($retries -eq 1) {
        Write-Host 'Server returned busy exception.  Retry up to 10 times.'
      }
      Write-Host "Retry # $retries :  Wait $(65 * $retries) seconds" -NoNewline
      Start-Sleep -Seconds (65 * $retries)
      Write-Host "    Retrying..."
    }
    Catch {
      Write-Host $response.ErrorMessage
      Return $response
    }
  } While ($retries -lt 11)

  throw 'Failed to get results after 10 retries.'
}

#Collects results by performing paged searches for given batch of mailboxes
#  Returns full set of PreviewItems along with MailboxStats
function Start-BatchCustomEdiscoverySearch {
  Param ($service, $KQL,
    #Prevent max allowable searchable mailboxes
    [ValidateScript({$_.Count -lt '15000'})]
    $Mailboxes,
    [int]$PageSize = '1000'
  )
  $Previews = New-Object System.Collections.ArrayList
  $Stats = New-Object System.Collections.ArrayList
  Do {
    $splGetSearchResults = @{
      service = $service
      KQL = $kql
      Mailboxes = $Mailboxes
      PageSize = $PageSize
      PageItemReference = $PageItemReference
    }
    $result = Get-EdiscoveryPagedSearchResults @splGetSearchResults

    if ($result.PreviewItems) {
      #Set the new page token
      $PageItemReference = $result.PreviewItems[-1].SortValue
      $Previews.AddRange($result.PreviewItems) | Out-Null
    }
    if ($result.MailboxStats) {
      $PositiveStats = $result.MailboxStats |
        where {$_.ItemCount -gt 0}
      if ($PositiveStats.count -gt 1) {
        $Stats.AddRange($PositiveStats) | Out-Null
      }
      elseif ($PositiveStats.count -eq 1) {
        $Stats.Add($PositiveStats) | Out-Null
      }
    }
  } While (
      #while pageitemcount = pagesize, keep going
      $result.PageItemCount -eq $PageSize
  )
  $hash = @{
    PreviewItems = $Previews
    MailboxStats = $Stats
  }
  New-Object psobject -Property $hash
}

#Returns impersonated O365 EWS service object
function New-ImpersonatedEwsO365Service {
  [CmdletBinding()]
  [Alias('Get-ImpersonatedEwsService')]
  Param (
    [Parameter(Mandatory=$true,
    ValueFromPipelineByPropertyName=$true,
    Position=0)]
    $Impersonate,
    [switch]$O365,
    $Credential
  )
  if ($O365) {
    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchVer)

    $ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress), $Impersonate
    $ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
    $service.ImpersonatedUserId = $ImpUserId

    $url = 'https://outlook.office365.com/EWS/Exchange.asmx'
    if (! $Credential) {
      $Credential = (Get-Credential -UserName (
        $env:UserName + '@domain.com') -m 'Credentials for O365 admin access'
      ).GetNetworkCredential()
    }
    $service.Credentials = New-Object System.Net.NetworkCredential -ArgumentList $Credential.UserName,
      $Credential.Password, $Credential.Domain
    $service.url = New-Object Uri($url)
  }
  else {
    $ExchVer = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
    $Service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($exchver)
    $ArgumentList = ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress),$Impersonate
    $ImpUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId -ArgumentList $ArgumentList
    $service.ImpersonatedUserId = $ImpUserId

    $service.UseDefaultCredentials = $true
    $service.AutodiscoverUrl($Impersonate)
  }
  #15 min timeout
  $service.Timeout = '900000'
  Return $service
}

#Binds to message item on o365/server
filter Get-EwsMailItem {
  Param (

   [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [Microsoft.Exchange.WebServices.Data.PreviewItemMailbox]
    $Mailbox,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [Microsoft.Exchange.WebServices.Data.ItemId]
    $Id,
    [PSCredential]
    $Credential,
    #increasing wait while retrying, multiples of 10 seconds
    [int] $Attempts = 5
  )
  #Declare properties we want returned for item
  $Subject = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject
  $DateTimeSent = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent
  $From = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender
  $ItemClass = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass
  $IdOnly =  [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
  $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
    $IdOnly, $DateTimeSent, $From, $Subject, $ItemClass
  )

  $splImpersonationParams = @{
    Impersonate = $Mailbox.PrimarySmtpAddress
    O365 = $true
    Credential = $Credential
  }
  $service = New-ImpersonatedEwsO365Service @splImpersonationParams

  #Retry routine
  $retries = 0
  Do {
    Try {
      #Bind to item, requesting only subset of values
      $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind(
        $service, $Id, $PropertySet
      )
      #let's tack on two values to be returned with EmailMessage obj
      $hash =  @{
        EmailMessage = $item
        Mailbox = $Mailbox
      }
      Return New-Object psobject -Property $hash
    }
    Catch {
      $retries++
      if ($retries -eq 1) {
        Write-Host "Error binding to item.  Retry up to $Attempts times."
      }
      Write-Host "Retry # $retries :  Wait $(10 * $retries) seconds" -NoNewline
      Start-Sleep -Seconds (10 * $retries)
      Write-Host "    Retrying..."
    }
  } While ($retries -lt $Attempts)
  Write-Host "Failed to bind to item  after $Attempts retries."
  Return
}

#If confirmed match, send Mailitem down pipeline for deletion
filter Confirm-eDiscoverySearchResult {
  Param (
    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
  [Microsoft.Exchange.WebServices.Data.EmailMessage]
  $EmailMessage,
  [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
  [Microsoft.Exchange.WebServices.Data.PreviewItemMailbox]
  $Mailbox,
    [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
   $ParentFolderName,
   $SenderAddresses, $RetentionAge = '-90', $ItemClass = 'IPM.Note'
  )
  #Add '*' to perform wildcard match
  $ItemClass = $ItemClass + '*'
  #Correct RetentionAge
  if ($RetentionAge -gt 0) {
    $RetentionAge = 0 - $RetentionAge
  }

    #Mailbox Item must meet 3 criteria below before we can hard delete item
  if ($EmailMessage.Sender.Address -notin $SenderAddresses) {
    $Confirmed = $false
  }
  elseif ($EmailMessage.DateTimeSent -ge (Get-Date).AddDays($RetentionAge)) {
    $Confirmed = $false
  }
  elseif ($EmailMessage.ItemClass -notlike $ItemClass) {
    $Confirmed = $false
  }
  else {
    $Confirmed = $true
  }

  # let's tack on two values to be returned with EmailMessage obj
  $rtnObject = @{
        Mailbox = $Mailbox
        EmailMessage = $EmailMessage
        ParentFolderName = $ParentFolderName
        MeetsCriteria = $true
  }
  if ($Confirmed) {
    Return New-Object psobject -Property $rtnObject
  }
  else {
    Write-Host 'Item found that does not meet search criteria'
    Write-Host "Mailbox: $($Mailbox.PrimarySmtpAddress)"
    Write-Host "Sender:  $($EmailMessage.Sender.Address)"
    Write-Host $EmailMessage.DateTimeSent
    Write-Host $EmailMessage.Id.UniqueId
    Write-Host "`r`n`r`n"

    $rtnObject.MeetsCriteria = $false
    Return New-Object psobject -Property $rtnObject
  }
}

#The function provides various options for item deletion
filter Remove-EwsMailItem {
  Param (
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [Microsoft.Exchange.WebServices.Data.EmailMessage]
    $EmailMessage,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    [Microsoft.Exchange.WebServices.Data.PreviewItemMailbox]
    $Mailbox,
    #Specify deletion mode
    [ValidateSet('HardDelete', 'SoftDelete', 'MoveToDeletedItems')]
    $DeleteMode = 'HardDelete',
    $Attempts = '3'
  )

  $Mode = [Microsoft.Exchange.WebServices.Data.DeleteMode]::$DeleteMode
  #Retry routine
  $retries = 0
  Do {
    Try {
      $EmailMessage.Delete($Mode)
      Write-Host "Deleted: $($EmailMessage.Id.UniqueId)"
      Return
    }
    # catch object not found, and bail
    Catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException]{
      Write-Host "Cannot Delete: Object no longer exists in folder:"
      Write-Host "Mailbox: $($Mailbox.PrimarySmtpAddress)"
      Write-Host "ParentId: $($item.ParentFolderId)"
      Write-Host "Id: $($item.Id) `r`n"
      Return
    }
    Catch {
      $retries++
      if ($retries -eq 1) {
        Write-Host "Error deleting item.  Retry up to $Attempts times."
      }
      Write-Host "Retry # $retries :  Wait $(10 * $retries) seconds" -NoNewline
      Start-Sleep -Seconds (10 * $retries)
      Write-Host "    Retrying..."
    }
  } While ($retries -lt $Attempts)
  throw "Failed to delete item after $Attempts retries."
}

filter Get-EwsMailItemWithParentFolderName {
  Param (

   [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    #[Microsoft.Exchange.WebServices.Data.PreviewItemMailbox]
    $Mailbox,
    [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
    #[Microsoft.Exchange.WebServices.Data.ItemId]
    $Id,
    [PSCredential]
    $Credential,
    #increasing wait while retrying, multiples of 10 seconds
    [int] $Attempts = 3
  )
  #Declare properties we want returned for item
  $Subject = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject
  $ParentFolderid = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ParentFolderId
  $DateTimeSent = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::DateTimeSent
  $From = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Sender
  $ItemClass = [Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass
  $BasePropertyIdOnly =  [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly
  $PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
    $BasePropertyIdOnly, $DateTimeSent, $From, $Subject, $ItemClass,
    $ParentFolderid
  )

  $splImpersonationParams = @{
    Impersonate = $Mailbox.PrimarySmtpAddress
    O365 = $true
    Credential = $Credential
  }
  $service = New-ImpersonatedEwsO365Service @splImpersonationParams

  #Retry routine
  $retries = 0
  Do {
    Try {
      #Bind to item, requesting only subset of values
      $item = [Microsoft.Exchange.WebServices.Data.Item]::Bind(
        $service, $Id, $PropertySet
      )
      $Folder = Get-EwsFolderById -service $service -FolderId $item.ParentFolderId

      #let's tack on two values to be returned with EmailMessage obj
      Return New-Object psobject -Property @{
        EmailMessage = $item
        Mailbox = $Mailbox
        ParentFolderName = $Folder.DisplayName
      }
    }
    # catch object not found, and bail
    Catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException]{
      Write-Host "Cannot bind to object.  Object no longer exists in folder:"
      Write-Host "Mailbox: $($Mailbox.PrimarySmtpAddress)"
      Write-Host "ParentId: $($item.ParentFolderId)"
      Write-Host "Id: $($item.Id) `r`n"
      Return
    }
    # for all other exceptions, such as ServerBusy, let's retry
    Catch {
      $retries++
      if ($retries -eq 1) {
        Write-Host "Error binding to item.  Retry up to $Attempts times."
      }
      Write-Host "Retry # $retries :  Wait $(10 * $retries) seconds" -NoNewline
      Start-Sleep -Seconds (10 * $retries)
      Write-Host "    Retrying..."
    }
  } While ($retries -lt $Attempts)
  Write-Host "Failed to bind to item  after $Attempts retries."
  Return
}

filter Get-EwsFolderById {
  Param (
    $FolderId,
    $service,
    $Attempts = 3
  )
  $retries = 0
  Do {
    Try {
      $fId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($FolderId.UniqueId)
      $folderObj = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$fId)
      Return $folderObj
    }
    # catch object not found, and bail
    Catch [Microsoft.Exchange.WebServices.Data.ServiceResponseException]{
      Write-Host "Cannot bind to folder.  Folder not found:"
      Write-Host "Mailbox: $($service.ImpersonatedUserId.Id)"
      Write-Host "FolderId: $($FolderId.UniqueId) `r`n"
      Return
    }
    # for all other exceptions, such as ServerBusy, let's retry
    Catch {
      $retries++
      if ($retries -eq 1) {
        Write-Host "Error binding to item.  Retry up to $Attempts times."
      }
      Write-Host "Retry # $retries :  Wait $(10 * $retries) seconds" -NoNewline
      Start-Sleep -Seconds (10 * $retries)
      Write-Host "    Retrying..."
    }
  } While ($retries -lt $Attempts)

  Write-Host "Failed to bind to folder after $Attempts retries."
  Return
}
