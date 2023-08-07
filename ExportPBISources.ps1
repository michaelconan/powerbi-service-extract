<# 
    Author: Michael Conan
    Description: Script to export Power BI workspaces, reports, datasets, dataflows and datasources for indiviudal or organization
    Usage: 
    - Admins: .\ExportPBISources.ps1 -scope Organization
#>

# Defaults to individual
param ($scope="Individual")
Write-Output "Running with $scope scope"

$base = ""
if ($scope -eq "Organization") {
    $base="admin/"
}

# Module installations
foreach ($mod in @("MicrosoftPowerBIMgmt", "ImportExcel")) {
    $check = Get-InstalledModule -Name $mod
    if ($check -eq $null) {
        Install-Module -Name $mod -Scope CurrentUser
    } else {
        Import-Module $mod
    }
}

# Login
Login-PowerBI

# Setup Variables
$reports = @()
$datasets = @()
$dataflows = @()
$tables = @()
$setflowlinks = @()
$datasources = @()

if ($scope -eq "Organization") {
    $workspaces = Get-PowerBIWorkspace -Type Workspace -First 5000 -Scope $scope
} else {
    $workspaces = Get-PowerBIWorkspace -All -Scope $scope
}

Write-Output "$($workspaces.Length) workspaces identified"

if ($workspaces.Length -gt 0) {
    # Loop workspaces, datasets, sources
    foreach ($ws in $workspaces)
    {
        Write-Output "Processing workspace: $($ws.Id)"
        
        # Reports
        $rpts = Get-PowerBIReport -WorkspaceId $ws.Id -Scope $scope
        $rpts | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
        }
        $reports = $reports + $rpts
        
        # Datasets
        $dsets = Get-PowerBIDataset -WorkspaceId $ws.Id -Scope $scope
        $dsets | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
        }
        $datasets = $datasets + $dsets

        Write-Output "$($dsets.Length) datasets identified"

        # Dataset datasources
        foreach ($ds in $dsets)
        {
            Write-Output "Processing dataset: $($ds.Id)"

            # $tbls = Get-PowerBITable -DatasetId $ds.Id -WorkspaceId $ws.Id -Scope $scope
            # $tables = $tables + $tbls

            <#
            # Cmdlet version
            $dsources = Get-PowerBIDatasource -DatasetId $ds.Id -WorkspaceId $ws.Id -Scope $scope
            $dsources | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
                $_ | Add-Member -MemberType NoteProperty -Name "DatasetId" -Value $ds.Id
                $_ | Add-Member -MemberType NoteProperty -Name "DataflowId" -Value $null
                $conn = $_.connectionDetails | ConvertTo-Json -Compress
                $_ | Add-Member -MemberType NoteProperty -Name "Connection" -Value $conn
            }
            $datasources = $datasources + $dsources
            #>

            # REST version
            $dsresp = Invoke-PowerBIRestMethod -Url "$($base)/datasets/$($ds.Id)/datasources" -Method GET
            if ($dsresp -ne $null) {
                $dsources = $dsresp | ConvertFrom-Json
                $dsources = $dsources.value
                $dsources | ForEach-Object {
                    $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
                    $_ | Add-Member -MemberType NoteProperty -Name "DatasetId" -Value $ds.Id
                    $_ | Add-Member -MemberType NoteProperty -Name "DataflowId" -Value $null
                    if ($_.PSObject.Properties.Name -contains 'connectionDetails') {
                        $_.connectionDetails = $_.connectionDetails | ConvertTo-Json -Compress
                    }
                }            
                $datasources = $datasources + $dsources
            }
            
  
        }

        # Dataflows
        $dflows = Get-PowerBIDataflow -WorkspaceId $ws.Id -Scope $scope
        $dflows | ForEach-Object {
            $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
        }
        $dataflows = $dataflows + $dflows

        Write-Output "$($dflows.Length) dataflows identified"

        # Dataflow datasources
        foreach ($df in $dflows) {
            
            <#
            # Cmdlet version
            $dfsources = Get-PowerBIDataflowDatasource -DataflowId $df.Id -WorkspaceId $ws.Id -Scope $scope
            $dfsources | ForEach-Object {
                $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
                $_ | Add-Member -MemberType NoteProperty -Name "DatasetId" -Value $null
                $_ | Add-Member -MemberType NoteProperty -Name "DataflowId" -Value $df.id
                $conn = $_.connectionDetails | ConvertTo-Json -Compress
                $_ | Add-Member -MemberType NoteProperty -Name "Connection" -Value $conn
            }
            $datasources = $datasources + $dfsources
            #>
            
            # REST version
            $dfresp = Invoke-PowerBIRestMethod -Url "$($base)dataflows/$($df.Id)/datasources" -Method GET
            if ($dfresp -ne $null) {
                $dfsources = $dfresp | ConvertFrom-Json
                $dfsources = $dfsources.value
                $dfsources | ForEach-Object {
                    $_ | Add-Member -MemberType NoteProperty -Name "WorkspaceId" -Value $ws.Id
                    $_ | Add-Member -MemberType NoteProperty -Name "DatasetId" -Value $null
                    $_ | Add-Member -MemberType NoteProperty -Name "DataflowId" -Value $df.Id
                    if ($_.PSObject.Properties.Name -contains 'connectionDetails') {
                        $_.connectionDetails = $_.connectionDetails | ConvertTo-Json -Compress
                    }
                }
                $datasources = $datasources + $dfsources
            }

        }

        # Dataset / Dataflow linkage
        $sfresp = Invoke-PowerBIRestMethod -Url "$($base)/groups/$($ws.id)/datasets/upstreamDataflows" -Method GET
        if ($sfresp -ne $null) {
            $sflinks = $sfresp | ConvertFrom-Json
            $sflinks = $sflinks.value
            $setflowlinks = $setflowlinks + $sflinks
        }

    }

    # Consolidated data for JSON
    $consolidated = @{
            Workspaces = $workspaces
            Reports = $reports
            Datasets = $datasets
            Dataflows = $dataflows
            Datasources = $datasources
            SetFlowLinks = $setflowlinks
    }
    $consolidated | ConvertTo-Json | Out-File -FilePath PowerBISources.json

    # Export data
    Write-Output "Export totals:"
    $summary = @{
        Workspaces = $workspaces.Length
        Reports = $reports.Length
        Datasets = $datasets.Length
        Dataflows = $dataflows.Length
        Datasources = $datasources.Length
        SetFlowLinks = $setflowlinks.Length
    }
    $summary | Out-String

    Write-Output "Writing data to Excel file"
    $output = "PowerBISources.xlsx"
    $workspaces | Export-Excel $output -WorksheetName workspaces
    $reports  | Sort-Object -Property WorkspaceId,Id -Unique | Export-Excel $output -WorksheetName reports
    $datasets  | Sort-Object -Property WorkspaceId,Id -Unique | Export-Excel $output -WorksheetName datasets
    # $tables  | Sort-Object -Property WorkspaceId,DatasetId,TableId -Unique | Export-Excel $output -WorksheetName tables
    $datasources | Sort-Object -Property WorkspaceId,DatasetId,DataflowId,DatasourceId -Unique | Export-Excel $output -WorksheetName datasources
    if ($dataflows.Length -gt 0) {
        $dataflows | Sort-Object -Property WorkspaceId,Id -Unique | Export-Excel $output -WorksheetName dataflows
        $setflowlinks | Export-Excel $output -WorksheetName datasetflowlinks
    }
}
else {
    Write-Output "No workspaces identified, ending process"
}
