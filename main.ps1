# install required modules
if( -not (Get-Module -ListAvailable -Name PSWriteHTML) ){
  Write-host "Installing PSWriteHTML Module"
  Install-Module PSWriteHTML -Force -AllowClobber -Scope CurrentUser
}
if( -not (Get-Module -ListAvailable -Name ReportingServicesTools) ){
  Write-host "Installing Reporting Services Module"
  Install-Module ReportingServicesTools -Force -AllowClobber -Scope CurrentUser
}

# ssrs server info
$ssrs_server = "my-server"
$ssrs_uri = "http://my-server/ReportServer"
$ssrs_portal_uri = "http://my-server/Reports"
$ssrs_credential_path = ".\ssrs_cred.xml"

# folder paths
$reports_path = ".\Reports"
$reports_archived_path = Join-Path -Path $reports_path -ChildPath "Archived"
$datasets_path = ".\Datasets"
$datasets_archived_path = Join-Path -Path $datasets_path -ChildPath "Archived"

# create necessary folders
foreach($path in @($reports_path, $reports_archived_path, $datasets_path, $datasets_archived_path)){
  if( (Test-Path -Path $path) -eq $false){
    New-Item -ItemType Directory -Path $path -Force
    Write-host "Created $($path) Directory"
  }
}

# generate ssrs credential
if( (Test-Path -Path $ssrs_credential_path) -eq $false){
  Write-host "Generating SSRS Credential to $($ssrs_credential_path)"
  Get-Credential | Export-Clixml -Path $ssrs_credential_path
}
$ssrs_cred = Import-Clixml -Path $ssrs_credential_path


# establish connection to ssrs server
try{
  Connect-RsReportServer -ComputerName $ssrs_server -ReportServerUri $ssrs_uri -Credential $ssrs_cred
  Write-host -ForegroundColor Green "Connected to Report server: ( $($ssrs_server) ) Report Server URI: ( $($ssrs_uri) )"
}
catch{
  throw $_
}

# get all of the contents from the ssrs server
$ssrs_contents = Get-RsFolderContent -RsFolder "/" -Recurse | select-object -property Name,Path,TypeName

# filter to only get the reports
$ssrs_reports = $ssrs_contents | where-object {$_.TypeName -eq "Report"} 

# go through each of the reports
foreach($report in $ssrs_reports){
  Write-Progress -Activity "Loop" -Status " $( [math]::round(([array]::IndexOf($ssrs_reports, $report) / $ssrs_reports.length) * 100))% COMPLETE" -PercentComplete $( ([array]::IndexOf($ssrs_reports, $report) / $ssrs_reports.length) * 100)
  Write-host "=== Report ===="
  Write-host "Name:       $($report.Name)"
  Write-host "SSRS Path:  $($report.Path)"

  # datasource info for the report
  $datasource = Get-RsItemDataSource -RsItem $report.path

  # collect the datasource information
  $name = ""
  $reference = ""
  if($datasource.length -gt 1){
    [string]$name = $datasource.Name -join "; "
    foreach($x in $datasource){
      [string]$reference += $x.Item.Reference + "; "
    }
  }
  else{
    [string]$name = $datasource.Name
    [string]$reference = $datasource.Item.Reference
  }

  # Add the datasource information to the report
  $report = $report | Add-Member -MemberType NoteProperty -Name "DataSourceName" -Value $name -PassThru | Add-Member -MemberType NoteProperty -Name "DataSourceReference" -Value $reference -PassThru

  # Download the report to /Reports
  Out-RsRestCatalogItem -RsItem $report.Path -Destination $reports_path -ReportPortalUri $ssrs_portal_uri

  # get the recently downloaded report file
  $file_name = Get-ChildItem -Path $reports_path -Name -File
  write-host -ForegroundColor Magenta "FETCHED Report:   $(Join-Path -Path $reports_path -ChildPath $file_name)"

  # get the file in xml format
  [xml]$xml = Get-Content -Path (Join-Path -Path $reports_path -ChildPath $file_name)

  # get all of the datasets for the report
  $datasets = $xml.Report.DataSets.Dataset
  $count = 0
  foreach($dataset in $datasets){
    # uses embedded query
    if($dataset.Query){
      $report = $report | Add-Member -MemberType NoteProperty -Name "Dataset $($count)" -Value ($dataset.Query.CommandText) -PassThru
      $str_result = ""
      foreach($field in $dataset.fields.field){
        $str_result += "$($field.Name) ($($field.TypeName)) ;"
      }
      $report = $report | Add-Member -MemberType NoteProperty -Name "Dataset $($count) Fields" -Value $str_result -PassThru
    }
    # uses a shared dataset
    elseif($dataset.SharedDataSet){
      try{
        Out-RsCatalogItem -RsItem $dataset.SharedDataSet.SharedDataSetReference -Destination $datasets_path
        $ds_file_name = Get-ChildItem -Path $datasets_path -Name -File
        Write-host -ForegroundColor Magenta "FETCHED Dataset:  $(Join-Path -Path $datasets_path -ChildPath $ds_file_name)"
        [xml]$ds_xml = Get-Content -Path (Join-Path -Path $datasets_path -ChildPath $ds_file_name)
        $report = $report | Add-Member -MemberType NoteProperty -Name "Dataset $($count)" -Value ($ds_xml.SharedDataSet.Dataset.Query.CommandText) -PassThru
        $str_result = ""
        foreach($field in $ds_xml.SharedDataSet.DataSet.Fields.Field){
          $str_result += "$($field.Name) ($($field.TypeName)) ;"
        }
        $report = $report | Add-Member -MemberType NoteProperty -Name "Dataset $($count) Fields" -Value $str_result -PassThru
        Move-Item -Path (Join-Path -Path $datasets_path -ChildPath $ds_file_name) -Destination $datasets_archived_path -Force
        write-host -ForegroundColor DarkBlue "MOVED Dataset:    $(Join-Path -Path $datasets_path -ChildPath $ds_file_name) ====> $($datasets_archived_path)"
      }
      catch{
        write-Error "Unable to get dataset for $($dataset.Name)"
      }
    }
    $count += 1
  }
  Move-Item -Path (Join-Path -Path $reports_path -ChildPath $file_name) -Destination $reports_archived_path -Force
  write-host -ForegroundColor DarkBlue "MOVED Report:     $(Join-Path -Path $reports_path -ChildPath $file_name) ====> $($reports_archived_path)"

  # Add the subscription info 
  $subscriptions_for_report = Get-RsSubscription -RsItem $report.Path
  $sub_body = ""
  foreach($sub in $subscriptions_for_report){
    $sub_body += "`n======================================"
    $sub_body += $sub | Select-Object -Property Description,Status,LastExecuted,@{N="DeliverySettings";E={$_.DeliverySettings | select-object -ExpandProperty ParameterValues | Format-Table -HideTableHeaders -Property Name,Value | out-string}} | format-list | out-string
    [xml]$xml = $sub.MatchData
    $properties = $xml.ScheduleDefinition | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name | Where-object {$_ -notin @('xsd', 'xsi')}
    foreach($prop in $properties){
      $sub_body += $xml.ScheduleDefinition.$prop | Select-object -Property ($xml.ScheduleDefinition.$prop | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name | Where-Object {$_ -notin 'xmlns'}) | format-List | out-string
    }
    $sub_body += "======================================`n`n"
  }
  $report = $report | Add-Member -MemberType NoteProperty -Name "Subscriptions" -Value $sub_body -PassThru
}

# generate webpage
Dashboard -Online -FilePath "output.html" -ShowHTML -Author "Tresten Pool" -TitleText "SSRS Reports"{
  Tab -Name "Detailed"{
    Panel{
      Table -DataTable $ssrs_reports -HideFooter -AllProperties -PreContent {"<h1>SSRS Reports From $($ssrs_server)</h1>"}
    }
  }
  Tab -Name "Datasource Information minified" {
    Panel {
      Table -DataTable $ssrs_reports -HideFooter -PreContent {"<h1>SSRS Reports From $($ssrs_server)</h1>"} -IncludeProperty "Name","Path","DataSourceName","DataSourceReference"
    }
  }
  Tab -Name "Subscriptions"{
    Panel {
      Table -DataTable ($ssrs_reports | Select-Object -Property Name,Path,Subscriptions) -HideFooter
    }
  }
}