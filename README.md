# Active-Directory-Automation
Active Directory Automation program using PowerShell

1. Show-fileOpenStatus: Check if input & output Excel files are already open
2. Add-NewADUserToOU: Read input Excel file (inputNewADUser.xlsx), check if users already exist in the system, add new AD Users to Active Directory system using New-ADUser, copy membership from default accounts
3. Add-UsersToFinalExcelFile: Log AD Account Creation information in output Excel file (NewADUserLog.xlsx)
4. Write-ExecutionSummary: Write execution summary in ADAutomationExecutionLog.txt

# License
This project is licensed under the MIT License - see the LICENSE file for details
