<#
.SYNOPSIS
This is a simple Powershell script to get SMB-opened files and folders
 and force close existing connections

.EXAMPLE
Get list of opened files and close them immediately:
    GetSMBOpenedFiles.ps1 -FolderPath 'C:\SharedFiles\Subfolder' -ForceClose

Get list of opened files without closing sessions:
    GetSMBOpenedFiles.ps1 -FolderPath 'C:\SharedFiles\Subfolder' -ListFiles

Get only count of opened files:
    GetSMBOpenedFiles.ps1 'C:\SharedFiles\Subfolder'

.NOTES
Don't forget about quotes if your path contains spaces!

.AUTHOR
Anton_Samsonov1@epam.com
12.02.2020
#>
Param (
    [Parameter (Mandatory = $true, Position = 1)]
    [string]$FolderPath,
    [Parameter (Mandatory = $false)]
    [switch]$ListFiles,
    [Parameter (Mandatory = $false)]
    [switch]$ForceClose
)

### This function use openfiles.exe tool to list opened files instead unavailable Get-SMBFiles commandlet on a server 2008 systems
function Get-FilesOpened {
    $openfiles = openfiles.exe /query /fo csv /V
    $openfiles | ForEach-Object {
        $line = $_
        if ($line -match '","') { $line }
    } | ConvertFrom-Csv
}

try {
    ### filter input values, make a search pattern from folderPath
    if ($FolderPath -ne '*') {
        if ($FolderPath[$FolderPath.Length - 1] -eq '\') {
            $FolderPath = $FolderPath + '*' 
        }
        else { 
            $FolderPath = $FolderPath + '\*' 
        }
    }

    ### Getting list of opened files and folders
    [array]$openedFiles = Get-FilesOpened | Where-Object { $_.'Open File (Path\executable)' -like $FolderPath }
    ### Showing result
    if ($ListFiles) {
        $openedFiles | Format-Table | Out-String | Write-Host -ForegroundColor Green 
    }
    "Summary $($openedFiles.count) connections to $FolderPath" | Write-Host -ForegroundColor Yellow
}
catch {
    $Error[0] | Write-Host -ForegroundColor Red
}
### Trying to close opened files and folders
if ($openedFiles.count -gt 0) {
    if ($ForceClose) {
        "Trying to close opened connections" | Write-Host -ForegroundColor Gray
        foreach ($fileSession in $openedFiles) {
            try {
                $null = & net file $fileSession.ID /close
                "$fileSession.ID sucsesfully closed" | Write-Host -ForegroundColor DarkYellow
            }
            catch {
                "Failed to close $fileSession.ID" | Write-Host -ForegroundColor DarkRed
            }
        }
    }
}
