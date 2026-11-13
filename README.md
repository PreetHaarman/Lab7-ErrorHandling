# Lab 7 – Error Handling

**Test-CriticalProcess**

Student: Harmanpreet Kaur

## Introduction

For this assignment, I wrote a PowerShell function called Test-CriticalProcess. It tests the processes lsass, dns, spoolsv and w3wp running on the local computer. If a process is detected, a green success message is displayed in the function. If the process is not identified, it will show a warning and create a text log for the missing process. The function generates a custom PowerShell object for each process, writes the results to an Excel workbook, and presents a summary of the number of processes checked, found, and not found. It also features parameters with validation and help with comments.

## Files

- `Test-CriticalProcess.ps1` – Script file
- `MissingProcesses.txt` – Log file
- `ProcessReport.xlsx` – Excel report
- `LAB_7_HarmanpreetKaur_ErrorHandling.docx` – Lab report

## Parameters

- `ProcessName` – One or more process names to check. The default names are lsass, dns, spoolsv, and w3wp.
- `LogPath` – The path of the text file used to record missing processes. The default is `.\MissingProcesses.txt`.
- `ExcelPath` – The path of the Excel report. The default is `.\ProcessReport.xlsx`.

## Requirements

The ImportExcel module is required to create the Excel report.

## How to Run

Test-CriticalProcess -Verbose

Test-CriticalProcess -ProcessName "notepad","lsass" -Verbose

Get-Help Test-CriticalProcess -Full

## Results

lsass and spoolsv are Found, and dns and w3wp are Missing.

## Limitation

This function checks processes on the local computer only.
PowerShell function Test-CriticalProcess with error handling, logging, Excel export and summary
