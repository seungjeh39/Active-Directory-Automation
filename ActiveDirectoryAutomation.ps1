# Active Directory Automation Project - Seungjeh Lee



# Check if necessary excel files are already open
Function Show-fileOpenStatus
{

    param ($File, $Path)

    $fileName = "$Path\$File.xlsx"
    $file = New-Object -TypeName System.IO.FileInfo -ArgumentList $fileName
    $ErrorActionPreference = "SilentlyContinue"
    [System.IO.FileStream] $fs = $file.OpenWrite();
    if (!$?) {
        return $false
    }
    else {
        $fs.Dispose()
        return $true
    }

}



# Create new AD Users and add them to Organizational Units (OU)
Function Add-NewADUserToOU
{

    param ($File, $Path)

    $ExcelObj = New-Object -comobject Excel.Application
    $ExcelObj.visible=$false
    $ExcelWorkBook = $ExcelObj.Workbooks.Open("$Path\$File.xlsx")
    $ExcelWorkSheet = $ExcelWorkBook.Sheets.Item(1)
    
    # Find # of rows and columns used
    $RowCount = 1
    while($null -ne $ExcelWorkSheet.Columns.Item(1).Rows.Item($RowCount).Value())
    {
        $RowCount++
    }
    $RowCount = $RowCount - 1
    $ColCount = 1
    while($null -ne $ExcelWorkSheet.Columns.Item($ColCount).Rows.Item(1).Value())
    {
        $ColCount++
    }
    $ColCount = $ColCount - 1

    # Define columns
    if($ColCount -eq 6)
    {
        $NameCol = [int]$ExcelWorkSheet.Range("A1:F1").find("Name").Column
        $EmployeeNumCol = [int]$ExcelWorkSheet.Range("A1:F1").find("Employee ID").Column
        $DepartmentCol = [int]$ExcelWorkSheet.Range("A1:F1").find("Department").Column
        $TitleCol = [int]$ExcelWorkSheet.Range("A1:F1").find("Title").Column
        $CommentCol = [int]$ExcelWorkSheet.Range("A1:F1").find("Comment").Column
        $sAMAccountNameCol = $EmployeeNumCol
    }
    else    # Unrecognized input format, write warning and close the excel file
    {
        Write-Warning "Unrecognized Input Format"

        $ExcelWorkBook.Save()
        $ExcelWorkBook.close($true)
        $ExcelObj.Quit()
        $ExcelObj = $ExcelWorkBook = $ExcelWorkSheet = $null
        [GC]::Collect()

        return $false
    }


    # Creating AD Users
    for($i=2; $i -le $RowCount; $i++)
    {   
        $fullName = $ExcelWorkSheet.Columns.Item($NameCol).Rows.Item($i).Text.Trim()
        $ndxOfLastName = $fullName.LastIndexOf(' ')
        $firstName = $fullName.split(' ')[0]
        $lastName = $fullName.Substring($ndxOfLastName+1, $fullName.length-$ndxOfLastName-1)
        
        $sAMAccountName = $ExcelWorkSheet.Columns.Item($sAMAccountNameCol).Rows.Item($i).Text.Trim()
        $employeeID = $ExcelWorkSheet.Columns.Item($EmployeeNumCol).Rows.Item($i).Text.Trim()
        $jobTitle = $ExcelWorkSheet.Columns.Item($TitleCol).Rows.Item($i).Text.Trim()

        # Choose Department if Department Column is empty 
        if($ExcelWorkSheet.Columns.Item($DepartmentCol).Rows.Item($i).Text.Trim() -like '')
        {
            $department = 'Asset Management','Board of Directors','Business Development','Corporate Communications',
            'Customer Service','Engineering','Finance','General Management','Human Resources','Information Technology',
            'Legal','Marketing','Operations','Product Management','Production','Project Management','Purchasing','Quality Assurance',
            'Risk Management','Sales','Technology' | Out-GridView -OutputMode Single -Title "Select the Department for $fullName"
        } else
        {
            $department = $ExcelWorkSheet.Columns.Item($DepartmentCol).Rows.Item($i).Text.Trim()
        }

        if($null -ne $CommentCol) {$description = $ExcelWorkSheet.Columns.Item($CommentCol).Rows.Item($i).Text.Trim()}
        else {$description = ''}
        $accountPW = 'password'     # default password
        $displayName = $fullName + ' (' + $department + ' - ' + $jobTitle + ')'
            

        # Choose a default account to copy (default accounts with pre-defined network domain accesses)
        $DefaultAccount = 'AssetMgmt','Director','BusinessDev','CorpComm','CustomerServ','Engineering','Finance','GenMgmt',
        'HR','IT','Legal','Marketing','Operations','ProdMgmt','Production','ProjMgmt','Purchasing','QualAssurance','RiskMgmt',
        'Sales','Technology' | Out-GridView -OutputMode Single -Title "Select a default account to copy for $displayName"
            

        $UserInstance = Get-ADUser -Filter {displayName -like $DefaultAccount} -Properties CannotChangePassword,Department,HomeDirectory,
        HomeDrive,PasswordExpired,PasswordNeverExpires,PasswordNotRequired
        #,ScriptPath,SmartcardLogonRequired


        # Choose an Organizationl Unit (OU) from the list (these OU's should be replaced with actual OU's)
        $OU = 'Admin','Contractors','CORP','Default Accounts','Engineering','Expats','Interns','IT','Managers',
        'Protected Accounts and Groups','Specialists','Supervisors','Technicians','Temporary Accounts','Test_Users',
        'Utility Accounts' | Out-GridView -OutputMode Single -Title "Select a group for $displayName"

        # Define new AD User parameters
        $NewADUser = @{
            Instance = $UserInstance
            SamAccountName = $sAMAccountName
            Name = $displayName
            GivenName = $firstName
            SurName = $lastName
            DisplayName = $displayName
            Office = "Employee ID: $employeeID"
            Title = $jobTitle
            Department = $department
            UserPrincipalName = $sAMAccountName
            Description = $description
            AccountPassword = (ConvertTo-SecureString $accountPW -AsPlainText -Force)
            # Path should be replaced
            Path = "OU=$OU,OU=Company Users,DC=Company Name,DC=com"
        }

        # Create new user
        if ($null -ne $OU)
        {
            New-ADUser @NewADUser
            Set-ADUser -Identity $sAMAccountName -ChangePasswordAtLogon $true
            # Copy Membership (network domain accesses)
            Get-ADUser -Filter {displayName -like $DefaultAccount} -Properties MemberOf | Select-Object -ExpandProperty MemberOf | Add-ADGroupMember -Members $sAMAccountName

            Write-Host $displayName' added to '$OU
                
            $global:newUserCount++
        }

    }

    # Save and close the Excel file
    $ExcelWorkBook.Save()
    $ExcelWorkBook.close($true)
    $ExcelObj.Quit()
    $ExcelObj = $ExcelWorkBook = $ExcelWorkSheet = $null
    [GC]::Collect()

    return $true

}



Function Add-UsersToFinalExcelFile
{
    param ($File, $Path)

    $ExcelObj1 = New-Object -comobject Excel.Application
    $ExcelObj1.visible=$false
    $ExcelWorkBook1 = $ExcelObj1.Workbooks.Open("$Path\$File.xlsx")
    $ExcelWorkSheet1 = $ExcelWorkBook1.Sheets.Item(1)

    # Find # of rows and columns used
    $RowCount1 = 1
    while($null -ne $ExcelWorkSheet1.Columns.Item(1).Rows.Item($RowCount1).Value())
    {
        $RowCount1++
    }
    $RowCount1 = $RowCount1 - 1
    $ColCount1 = 1
    while($null -ne $ExcelWorkSheet1.Columns.Item($ColCount1).Rows.Item(1).Value())
    {
        $ColCount1++
    }
    $ColCount1 = $ColCount1 - 1

    $ExcelObj2 = New-Object -comobject Excel.Application
    $ExcelObj2.visible=$true
    $ExcelWorkBook2 = $ExcelObj2.Workbooks.Open("$Path\NewADUserLog.xlsx")
    $ExcelWorkSheet2 = $ExcelWorkBook2.Sheets.Item(1)

    $ExcelObj1.Calculation = -4135
    $ExcelObj2.Calculation = -4135

    # For setti input
    if($ColCount1 -eq 8)
    {
        $NameCol = [int]$ExcelWorkSheet1.Range("A1:F1").find("Name").Column
        $EmployeeNumCol = [int]$ExcelWorkSheet1.Range("A1:F1").find("Employee ID").Column
        $TitleCol = [int]$ExcelWorkSheet1.Range("A1:F1").find("Title").Column
        $sAMAccountNameCol = $EmployeeNumCol
        $CommentCol = [int]$ExcelWorkSheet1.Range("A1:F1").find("Comment").Column
    }


    for($i=2; $i -le $RowCount1; $i++)
    {
        # Row to append the user info
        $RowCount2 = 1
        while($null -ne $ExcelWorkSheet2.Columns.Item(2).Rows.Item($RowCount2).Value())
        {
            $RowCount2++
        }

        # Append the user info to NewADUserLog.xlsx
        $UserObj = $(try {Get-ADUser $ExcelWorkSheet1.Columns.Item($sAMAccountNameCol).Rows.Item($i).Value()} catch {$null})
        if($null -ne $UserObj)
        {
            # Name
            $ExcelWorkSheet2.Columns.Item(2).Rows.Item($RowCount2).Value() = $ExcelWorkSheet1.Columns.Item($NameCol).Rows.Item($i).Value()
            # Employee ID
            $ExcelWorkSheet2.Columns.Item(3).Rows.Item($RowCount2).Value() = $ExcelWorkSheet1.Columns.Item($EmployeeNumCol).Rows.Item($i).Value()
            # Department
            $department = Get-ADUser -Identity $ExcelWorkSheet1.Columns.Item($sAMAccountNameCol).Rows.Item($i).Value() -Properties Department | Select-Object -ExpandProperty Department
            $ExcelWorkSheet2.Columns.Item(4).Rows.Item($RowCount2).Value() = Out-String -InputObject $department
            # Title
            $ExcelWorkSheet2.Columns.Item(5).Rows.Item($RowCount2).Value() = $ExcelWorkSheet1.Columns.Item($TitleCol).Rows.Item($i).Value()
            # Password
            $ExcelWorkSheet2.Columns.Item(7).Rows.Item($RowCount2).Value() = 'password'
            # Creation Date
            $whenCreated = Get-ADUser -Identity $ExcelWorkSheet1.Columns.Item($sAMAccountNameCol).Rows.Item($i).Value() -Properties whenCreated | Select-Object -ExpandProperty whenCreated
            $whenCreatedStr = Out-String -InputObject $whenCreated
            $start = $whenCreatedStr.IndexOf(' ') + 1
            $end = $whenCreatedStr.LastIndexOf(',') + 6
            $ExcelWorkSheet2.Columns.Item(8).Rows.Item($RowCount2).Value() = $whenCreatedStr.Substring($start, $end - $start)
            # Comment
            $ExcelWorkSheet2.Columns.Item(6).Rows.Item($RowCount2).Value() = $ExcelWorkSheet1.Columns.Item($CommentCol).Rows.Item($i).Value()
        }

    }
    
    $ExcelObj1.Calculation = -4105
    $ExcelObj2.Calculation = -4105

    # Save and close the Excel file
    $ExcelWorkBook1.Save()
    $ExcelWorkBook1.close($true)
    $ExcelWorkBook2.Save()
    $ExcelWorkBook2.close($true)

    $ExcelObj1.Quit()
    $ExcelObj2.Quit()
    $ExcelObj1 = $ExcelWorkBook1 = $ExcelWorkSheet1 = $null
    [GC]::Collect()
    $ExcelObj2 = $ExcelWorkBook2 = $ExcelWorkSheet2 = $null
    [GC]::Collect()
    
}



# Script execution log
Function Write-ExecutionSummary
{
    param ($Path)

    $date = (Get-Date).ToString()
    $user = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    $text = "$date $user - $global:newUserCount Users created"
    $text | Add-Content $Path'\ADAutomationExecutionLog.txt'

}



# My directory
$Path = 'C:\Users\Documents\ActiveDirectoryManagement'
# Input Excel File
$FileName = 'inputNewADUser'
# Count new users
$newUserCount = 0


$PatF1 = "$Path\$FileName.xlsx"
$inputDateModified = [datetime](Get-ItemProperty -Path $PatF1 -Name LastWriteTime).LastWriteTime

$Path2 = "$Path\ADAutomationExecutionLog.txt"
$programDateModified = [datetime](Get-ItemProperty -Path $Path2 -Name LastWriteTime).LastWriteTime

$timeSpan = New-TimeSpan -Start $inputDateModified -End $programDateModified

# if inputNewADUser.xlsx hasn't been updated since last execution, write a warning
if($timeSpan.ToString().Substring(0,1) -ne '-')
{
    Write-Warning "inputNewADUser.xlsx has not been updated since last execution. Please check the input file again."
} else
{
    # Check if the excel files are already open
    if (Show-fileOpenStatus -File $FileName -Path $Path)
    {
        if(Show-fileOpenStatus -File "NewADUserLog" -Path $Path)
        {
        
            # Load the Active Directory Module
            Import-Module -Name ActiveDirectory

            # Add new users from Excel file to given OU
            if(Add-NewADUserToOU -File $FileName -Path $Path)
            {
                # Add created users to NewADUserLog.xlsx
                Add-UsersToFinalExcelFile -File $FileName -Path $Path
            }

            # Log Execution History
            Write-ExecutionSummary -Path $Path

        } else
        {
            Write-Warning "Please close NewADUserLog.xlsx and try again"
        }
    } else
    {
        Write-Warning "Please close $FileName.xlsx and try again"
    }

}

# Execution completed
Read-Host -Prompt "Press Enter to exit"


