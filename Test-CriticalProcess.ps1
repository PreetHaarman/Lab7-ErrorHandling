<#
.SYNOPSIS
    Checks whether critical processes are running.

.DESCRIPTION
    Checks each process name given. Shows a success message if the process
    is running and a warning if it is not. Missing processes are written to
    a log file. A custom object is created for each process, the results are
    exported to Excel, and a summary is shown at the end.

.PARAMETER ProcessName
    One or more process names to check.
    Default: lsass, dns, spoolsv, w3wp.

.PARAMETER LogPath
    Path of the log file for missing processes.
    Default: .\MissingProcesses.txt

.PARAMETER ExcelPath
    Path of the Excel report.
    Default: .\ProcessReport.xlsx

.EXAMPLE
    Test-CriticalProcess -Verbose

.EXAMPLE
    Test-CriticalProcess -ProcessName "notepad","lsass" -Verbose

.NOTES
    Author: Harmanpreet Kaur
    Created: 2026-11-10
    The ImportExcel module is needed for the Excel report.
#>

#Parameters
# No script-level parameters are required.
# The function parameters are defined inside Test-CriticalProcess.

#Functions
function Test-CriticalProcess {

    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string[]]$ProcessName = @("lsass", "dns", "spoolsv", "w3wp"),

        [ValidateNotNullOrEmpty()]
        [string]$LogPath = ".\MissingProcesses.txt",

        [ValidateNotNullOrEmpty()]
        [string]$ExcelPath = ".\ProcessReport.xlsx"
    )

    $Results = @()

    foreach ($Name in $ProcessName) {

        Write-Verbose "Checking process $Name"

        try {
            # -ErrorAction Stop makes the error terminating so catch can handle it.
            $Process = Get-Process -Name $Name -ErrorAction Stop

            $Obj = [PSCustomObject]@{
                ProcessName = $Name
                Status      = "Found"
                ProcessId   = ($Process.Id -join ", ")
                CheckedTime = Get-Date
            }

            Write-Host "SUCCESS: $Name is running." -ForegroundColor Green
        }
        catch [Microsoft.PowerShell.Commands.ProcessCommandException] {
            # Process not found.
            $Obj = [PSCustomObject]@{
                ProcessName = $Name
                Status      = "Missing"
                ProcessId   = ""
                CheckedTime = Get-Date
            }

            Write-Warning "$Name was not found."
            Add-Content -Path $LogPath -Value "$(Get-Date) - $Name not found"
        }
        catch {
            # Any other process-checking error.
            $Obj = [PSCustomObject]@{
                ProcessName = $Name
                Status      = "Error"
                ProcessId   = ""
                CheckedTime = Get-Date
            }

            Write-Host "ERROR checking $Name : $($PSItem.Exception.Message)" -ForegroundColor Red
        }

        $Results += $Obj
    }

    # Export the results to Excel.
    try {
        Import-Module ImportExcel -ErrorAction Stop
        $Results | Export-Excel -Path $ExcelPath -AutoSize -ErrorAction Stop
    }
    catch {
        Write-Warning "Excel export failed: $($PSItem.Exception.Message)"
    }

    # Count the results for the summary.
    $FoundCount = ($Results | Where-Object Status -eq "Found").Count
    $MissingCount = ($Results | Where-Object Status -eq "Missing").Count

    # Display the colour-coded summary.
    Write-Host ""
    Write-Host "Summary" -ForegroundColor Cyan
    Write-Host "Processes checked : $($Results.Count)" -ForegroundColor White
    Write-Host "Processes found   : $FoundCount" -ForegroundColor Green
    Write-Host "Processes missing : $MissingCount" -ForegroundColor Yellow
}

#Execution Code
Test-CriticalProcess -Verbose
