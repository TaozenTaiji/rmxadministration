#Install-Module -Name InvokePsExec
#/Install-Module -Name CredentialManager

function Get-SqlConnectionString(){
    return "Data Source=tcp:rmxprod.database.windows.net,1433;Initial Catalog=rhythmedix;Authentication=Active Directory Integrated;"
  }

function Get-TolmanSqlConnectionString(){
    return "Data Source=tcp:rmxprod.database.windows.net,1433;Initial Catalog=Tolman;Authentication=Active Directory Integrated;"
  }
  
  Function Add-RhythmstarUser{
    [CmdLetBinding()]
      Param(
      [Parameter(Mandatory=$True)]$FullName,
      [Parameter(Mandatory=$false)]$Portal
      )
      
      if(!($portal))
      {
          $portal = read-host -prompt 'Which portal? RMX, Demo, or Tolman'
      }
      
      $SqlConnection = New-Object System.Data.SqlClient.SqlConnection

      switch($Portal)
      {
          'RMX'
          {
            $proceed = read-host -prompt "Add: $FullName to the Clinical Rhythmstar Portal? Y/N"
            $SqlConnection.ConnectionString = Get-SqlConnectionString
          }
          'Demo'
          {
            $proceed = read-host -prompt "Add: $FullName to the Demo Rhythmstar Portal? Y/N"
            $SqlConnection.ConnectionString = Get-DemoSqlConnectionString
          }
          'Tolman'
          {
            $proceed = read-host -prompt "Add: $FullName to the Tolman Rhythmstar Portal? Y/N"
            $SqlConnection.ConnectionString = Get-TolmanSqlConnectionString
          }
      }
    
      if($proceed -like 'Y')
      {
        $FirstInitial = $FullName.Substring(0,1)
        $FirstName, $LastName = $FullName -split "\s", 2
        $accountName = $FirstInitial + $LastName #login name
        $UPN = $accountName.ToLower() + "@rhythmedix.com" #userprincipalname
          $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
          $SqlCmd.CommandText = "dbo.spSystemCreateuser"  ## this is the stored proc name 
          $SqlCmd.Connection = $SqlConnection  
          $SqlCmd.CommandType = [System.Data.CommandType]::StoredProcedure  ## enum that specifies we are calling a SPROC
          $param1=$SqlCmd.Parameters.Add("@USERNAME" , [System.Data.SqlDbType]::VarChar)
              $param1.Value = $UPN 
          $param2=$SqlCmd.Parameters.Add("@FirstName" , [System.Data.SqlDbType]::VarChar)
              $param2.Value = $FirstName
          $param3=$SqlCmd.Parameters.Add("@LastName" , [System.Data.SqlDbType]::VarChar)
              $param3.Value = $LastName 
  
          $SqlConnection.Open()
          $result = $SqlCmd.ExecuteNonQuery() 
          Write-output "result=$result" 
          $SqlConnection.Close()
        }
  }
  
  Function Add-IPWhitelist{
    [CmdLetBinding()]
      Param(
      [Parameter(Mandatory=$True)]$UPN,
      [Parameter(Mandatory=$true)]$IP
      )
      
    
      $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
      $SqlConnection.ConnectionString = Get-SqlConnectionString
          $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
          $SqlCmd.CommandText = "dbo.spSystemUpdateFirewall"  ## this is the stored proc name 
          $SqlCmd.Connection = $SqlConnection  
          $SqlCmd.CommandType = [System.Data.CommandType]::StoredProcedure  ## enum that specifies we are calling a SPROC
          #SP format exec [spSystemUpdateFirewall] @UserName = 'vzobrak@rhythmedix.com', @ip = '1.10.1.15'
          $param1=$SqlCmd.Parameters.Add("@USERNAME" , [System.Data.SqlDbType]::VarChar)
              $param1.Value = $UPN 
          $param2=$SqlCmd.Parameters.Add("@IP" , [System.Data.SqlDbType]::VarChar)
              $param2.Value = $IP
          $SqlConnection.Open()
          $result = $SqlCmd.ExecuteNonQuery() 
          Write-output "result=$result" 
          $SqlConnection.Close()
        
  }
  function Disable-WVDSessionHost{
    [CmdLetBinding()]
    param()
    $sessionhost = read-host -prompt "Which session host? Enter full name: SessionHost.rhythmedix.com"
 
    $hostpool = "RemoteReview_HostPool"
    $tenantname = "RhythMedix Remote Review"
     Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
     Set-RdsSessionHost -TenantName $tenantname -HostPoolName $hostpool -Name $sessionhost -AllowNewSession:$false
  }
  function Enable-WVDSessionHost{
    [CmdLetBinding()]
    param()
    $sessionhost = read-host -prompt "Which session host? Enter full name: SessionHost.rhythmedix.com"
    $hostpool = "RemoteReview_HostPool"
    $tenantname = "RhythMedix Remote Review"
     Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
     Set-RdsSessionHost -TenantName $tenantname -HostPoolName $hostpool -Name $sessionhost -AllowNewSession:$true
  }

  function Get-WVDSession{
    [CmdLetBinding()]
    param(
        [Parameter(Mandatory=$False)]$HostName
    )
    $hostpool = "RemoteReview_HostPool"
    $tenantname = "RhythMedix Remote Review"
     Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
    if($null -ne $HostName)
    {
        Get-RdsUserSession -TenantName $tenantname -HostPoolName $hostpool | where-object { $_.SessionHostName -eq $hostname} | out-host
    }
    else 
    {
       Get-RdsUserSession -TenantName $tenantname -HostPoolName $hostpool  | out-host
    }

  }
  function Add-NewUser{
    [CmdLetBinding()]
      param(
      [Parameter(Mandatory=$True)]$GivenName,
      [Parameter(Mandatory=$True)]$SurName,
      [Parameter(Mandatory=$False)]$suffix="",
      [Parameter(Mandatory=$True)]$Title,
      [Parameter(Mandatory=$False)]$Department,
      [Parameter(Mandatory=$False)]$Manager,
      [Parameter(Mandatory=$False)]$Ladies,
      [Parameter(Mandatory=$False)]$Rhythmstar,
      [Parameter(Mandatory=$False)]$Location,
      [Parameter(Mandatory=$False)]$Remote
      )

      
      do
      {
      $Ladies = read-host -Prompt "Is this user a lady? Yes or No"
      <#
  
      Parameters to pass - [string]FullName , [String]Title, [string]Department (optional, unnecessary if title is explicitly listed here), 
      [string]manager (optional, if title isn't explicitly listed here)
          Departments: 
              Clinical
                  Arrhythmia Analyst
                  Sr. Arrhythmia Analyst
                  Holter technician
              Logistics
                  Product Distribution Specialist
              Clinical Administrators
                  Clinical Administrator
              All others - depends on title of person being hired, more custom titles
  
          Managers:
              Clinical (non-holter) - tcatling
              Clinical (holter) - ndemiranda
              Logistics - evalentine 
              Clinical Administrators - arichmann
              All others are rare
  
          Account name usually first letter of first name + lastname.
  
          Licenses:
              E3 - everyone going forward except for Logistics
              E1 - only logistics
              EMS E3 - holter, sales, VPs and anyone else who has a laptop/VPN
  
          To order/change license count - send email to Strong, Katrina <Katrina.Strong@softwareone.com> with number of licenses to add/remove. 
          We have our enterprise agreement with Microsoft through software one.
      #>
  
        $pattern = '[^a-zA-Z]'
        $samaccountname = ($givenname[0] + ($surname -replace $pattern, '') + $suffix).tolower()
        if ($suffix -ne "")
        {
              $displayname = $givenname + " $surname $suffix"
        }
        else {
            $displayname = $givenname + " $surname"
        }

        $FullName = $displayname
      
      
     
      switch($Title)
      {
          'Arrhythmia Analyst'
              {
                  $department = 'Clinical'
                  $Manager = 'tcatling'
                  break
              }
          'Sr. Arrhythmia Analyst'
              {
                  $department = 'Clinical'
                  $Manager = 'tcatling'
                  break
              }
          'Holter Technician'
              {
                  $department = 'Clinical'
                  $Manager = 'ndemiranda'
                  break
              }
          'Product Distribution Specialist'
              {
                  $department = 'Logistics'
                  $Manager = 'evalentine'
                  break
              }
          'Clinical Administrator'
              {
                  $department = 'Clinical Administrators'
                  $manager = 'arichmann'
                  break
              }
         
      }
      if($department -like 'Sales')
      {
        $Manager = 'kgartland'
      }
                  #'Clinical'
      #"Clinical" #Clinical, Logistics, IT, Sales, Payer Relations, Clinical Administrators, Engineering and Development
      #$title = "Arrhythmia Analyst" #"Product Distribution Specialist" #"Clinical Administrator" #"Arrhythmia Analyst" #"Holter Technician" #"Sr. Arrhythmia Analyst"
      #$manager = "tcatling" #don't need UPN here, just login name (first portion) #tcatling, evalentine, arichmann, ndemiranda
  
      
      $upn = $samaccountName + "@rhythmedix.com" #userprincipalname
     
      $tempPassword = convertto-securestring "Password1" -asplaintext -force
      if(!($Department))
      {
          $Department = read-host -prompt "What department is $displayname in:"
      }
      if(!($Manager))
      {
          $Manager= read-host -prompt "Who is $displayname's manager:(SamAccountName)"
      }
      $empNumber= " "
      $empNumber = read-host -prompt "Enter the Employee Number if present:"
      Write-Host "
                  User Name: $displayname
                  Title: $title
                  Department: $Department
                  Manager: $Manager
                  Email: $upn 
                  Employee Number: $empNumber"
        $continue = read-host -Prompt "Continue? Y/N"
    }while($continue -like 'N')
      $user = New-AdUser -Name $displayName -SamAccountName $samaccountName -AccountPassword $tempPassword -ChangePasswordAtLogon $true -Department $department -Title $title -DisplayName $displayName -EmailAddress $upn -GivenName $givenName -Surname $surname -Manager $manager -UserPrincipalName $upn -EmployeeID $empNumber -Enabled $true -PassThru
  
      #common for everyone
      #Add-AdGroupMember "All Employees" $user
      Add-AdGroupMember "Azure AD Sync" $user #required group to sync to cloud
      Add-AdGroupMember "Domain ADP Users" $user #group that allows SSO with ADP
      Add-AdGroupMember "Azure MFA Users" $user #group that turns on MFA Requirement
      
      Connect-MsolService -credential (get-storedcredential -target O365Admin)
    
      DO
      {		
              Sync-Azure
              Write-Host "." -NoNewline
              Start-Sleep -Seconds 10
      } Until (Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue)
      
     
      
      Set-MsolUser -UserPrincipalName $upn -UsageLocation "US"
      if ($Department -eq "Logistics")
      {
          Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:StandardPACK"
      }
      else
      {
        If ($title -eq 'Holter Technician')
        {
          $contractor = Read-Host "Is user a contractor? Y/N"
          if ($contractor -like 'y')
          {
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:StandardPACK"
          }
          else 
          {
            Add-AdGroupMember "VPN Users" $user
            Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:ENTERPRISEPACK"
          }
          
       }
      else 
      {
        Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:ENTERPRISEPACK"
      }
      }
      Sync-Azure
      Start-sleep -seconds 30
      Connect-EXO
      Add-O365GroupUser -GroupName "All Employees" -upn $upn
      if($location -ne "Leominster")
      {
          Add-AdGroupMember "Mount Laurel Office" $user
          $OnPrem = Read-host -Prompt "Will the user be in the Mt Laurel Office the majority of the time?"
          if ($OnPrem -like 'y')
          {
              Add-O365GroupUser -GroupName 'Mt Laurel Office' -upn $upn
          }
      }
      else
      {
          Add-AdGroupMember "Tolman Office" $user #if user is in MA (not hiring there right now)
      }
      
      if ($Ladies -like 'Y')
      {
          #ladies of rhythmedix - as it says
          Add-AdGroupMember "Ladies of Rhythmedix" $user
      }
  
      switch($Department)
      {
          'Clinical'
          {
              if ($Title -eq 'Holter Technician')
              {
                  #holter
                  Add-AdGroupMember "Holter Users" $user
                  Add-AdGroupMember "Self-Service Password Reset" $user
                  Add-AdGroupMember "Azure AD Domain Services" $user
                  Add-RhythmstarUser -Portal 'RMX' -FullName $FullName
                  Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:AAD_Premium"
              
                              
              }
              else
              {
                  $remote = read-host -prompt 'Is the tech a remote only worker?'
                  #monitoring center - arrhythmia analyst + sr. arrhythmia analyst
                 
                  Add-ADGroupMember "Clinical Schedule Viewers" $User #syncs with sharepoint online permissions
                  Add-AdGroupMember "Hourly Employees" $user
                  Add-AdGroupMember "Monitoring Techs" -Members $User
                  Add-RhythmstarUser -FullName $FullName
                  Add-WVDAppUser -user $samaccountName
                  if($remote -like 'Y')
                  {
                    Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:AAD_Premium"
                    Add-AdGroupMember "Self-Service Password Reset" $user
                  }
                  else {
                       #  Add-AdGroupMember "Monitoring" $user
                         Add-O365GroupUser -GroupName "RMX Monitoring" -upn $upn
                       if($location -eq "Leominster")
                        {
                            add-o365groupuser -GroupName "Leominster Monitoring" -upn $upn
                        }
                        else
                        {
                            Add-O365GroupUser -GroupName "Mt Laurel Monitoring" -upn $upn
                            Add-DistributionGroupMember -Identity "Verbal Orders" -member $upn
                        }
                    }
              }
          }
          'Clinical Administrators'
          {
              #clinical admin
                  Add-AdGroupMember "Hourly Employees" $user
                  Add-AdGroupMember "Ringcentral Softphone Users" $user
                  Add-AdGroupMember "ClinicalAdmins" $user
                  Add-RhythmstarUser -FullName $FullName
                  Add-O365GroupUser -GroupName 'Logistical Peeps' -upn $upn
                  Add-O365GroupUser -GroupName 'Clinical Admin' -upn $upn
                  Add-O365GroupUser -GroupName 'Customer Service - Comments' -upn $upn
                  Add-DistributionGroupMember -Identity "Customer Service" -member $upn
                  Add-WVDAppUser -user $samaccountName

          }
          'Logistics'
          {
              #logistics
                  Add-AdGroupMember "Hourly Employees" $user
                  Add-RhythmstarUser -FullName $FullName -Portal 'RMX'
                  Add-O365GroupUser -GroupName 'Logistical Peeps' -upn $upn
                  Add-AdGroupMember "Logistics" -Members $samaccountName
                  
          }
          'Sales'
          {
              #sales
                  Add-AdGroupMember "VPN Users" $user
                  Add-AdGroupMember "Self-Service Password Reset" $user
                  Add-RhythmstarUser -FullName $FullName -Portal 'RMX'
                  Add-RhythmstarUser -FullName $FullName -Portal 'Demo'
                  Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:EMS"
          }
          'Engineering'
          {
              #engineering
                  Add-AdGroupMember "Regulatory Medical Device" $user
                  Add-DistributionGroupMember -Identity "Regulatory Medical Device" -member $upn
                  Add-AdGroupMember "RPSS User" $user
          }
          'IT'
          {
              #IT
                  Add-O365GroupUser -GroupName "Rhythmedix IT" -upn $upn
                  Add-AdGroupMember "VPN Users" $user
                  Add-AdGroupMember "Self-Service Password Reset" $user
                  Add-RhythmstarUser -FullName $FullName
                  Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:EMS"
                 
                  If($Title -like "Developer")
                  {
                     Add-O365GroupUser -GroupName "Development" -Upn $upn
                     Add-O365GroupUser -GroupName "Dev Team" -Upn $upn
                  }
          }
  
      }
  
      
  
  
  
  
      
  }
  function Convert-ADUserToCloudOnly   {
    [CmdLetBinding()]
      param
      (
      [Parameter(mandatory=$true)]$accountname
      )
      $samaccountname = get-aduser -Identity $accountname
      $upn = $SAMaccountname.userprincipalname
      Remove-ADGroupMember -Identity "Azure AD Sync" -Members $samaccountname
  
      Sync-Azure
      Connect-MsolService -credential (get-storedcredential -target O365Admin)
      Get-MsolUser -UserPrincipalName $upn -ReturnDeletedUsers | Restore-MsolUser
      Get-MsolUser -UserPrincipalName $upn | Set-MsolUser -ImmutableId ""
      
      Get-Aduser -Identity $accountname | Remove-ADUser
  }  
  
  function Export-DLtoCSV{
    [CmdLetBinding()]
   param(
      [Parameter(Mandatory=$True)]$GroupName
      )
      $DGName = $GroupName
          Get-DistributionGroupMember -Identity $DGName | Select-Object Name, PrimarySMTPAddress |
          Export-CSV "C:\\Distribution-List-Members.csv" -NoTypeInformation -Encoding UTF8
  }
  
  Function Sync-Azure{
    [CmdLetBinding()]
    param()
      #PSFile version gives the on-screen feedback but requires ps1 file in the correct folder. The Command version doesn't give feedback but works regardless 
   # Invoke-PsExec -ComputerName galactica.rhythmedix.com -PSFile "C:\powershell tools\Sync_Azure.ps1" 
   Write-Host "Initializing Azure AD Delta Sync..." -ForegroundColor Yellow
    Invoke-PsExec -ComputerName galactica.rhythmedix.com -credential (get-storedcredential -target O365Admin)  -Command {
      
  
      Start-ADSyncSyncCycle -PolicyType Delta
  
      #Wait 10 seconds for the sync connector to wake up. 
      Start-Sleep -Seconds 10
  
      #Display a progress indicator and hold up the rest of the script while the sync completes.
      While(Get-ADSyncConnectorRunStatus)
      {
          Write-Host "." -NoNewline
          Start-Sleep -Seconds 10
      }
  
  Write-Host " | Complete!" -ForegroundColor Green } -IsPSCommand -IsLongPSCommand
    
      #Disconnect-EXO
  }

  function Restart-Holter  {
      [cmdletBinding()]
      param(
          [Parameter(Mandatory=$True)]$username
      )
      connect-AZaccount -credential (get-storedcredential -target O365Admin)
    get-azvm -resourcegroup 'RHythmedix-Infrastructure' | where-object {$_.Tags['tech'] -eq "$username@rhythmedix.com"} | Restart-AZVM
  }

  function Connect-O365Compliance{
    [CmdLetBinding()]
    param()
    if (!(get-pssession | where-object {$_.ConnectionURI -eq 'https://ps.compliance.protection.outlook.com/powershell-liveid/'}))
	{
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -credential (get-storedcredential -target O365Admin) -Authentication Basic -AllowRedirection
        start-sleep 5
        Import-PSSession $Session -DisableNameChecking -AllowClobber
       
    }
}

function Connect-EXO{
    [CmdLetBinding()]
    param()
    #$UserCredential = Get-StoredCredential -Target O365Admin
    if (!(get-pssession | where-object {$_.ConfigurationName -eq 'Microsoft.Exchange'}))
	{
        $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -credential (get-storedcredential -target O365Admin) -Authentication Basic  -AllowRedirection
        start-sleep 5
        Import-PSSession $ExoSession -DisableNameChecking -AllowClobber
       
	}

}

function Remove-Phishing{
    [CmdLetBinding()]
	param(
		[Parameter(Mandatory=$True)]$SearchName
	)
    Connect-O365Compliance
    New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete -confirm:$False
    Disconnect-O365Compliance
}

function Disconnect-EXO{
    [CmdLetBinding()]
    param()
    get-pssession | where-object {$_.ConfigurationName -eq 'Microsoft.Exchange'} | remove-pssession
	}

function Disconnect-O365Compliance{
    [CmdLetBinding()]
    param()
	get-pssession | where-object {$_.ComputerName -like '*compliance*'} | Remove-PSSession
	
}

function Add-ExoMailboxPermission{
    [CmdLetBinding()]
	param( 
		[string]$TargetMailboxOwner,
		[string]$User
		)
	connect-exo
    Add-MailboxPermission -Identity $TargetMailboxOwner -User $User -AccessRights FullAccess -InheritanceType All -AutoMapping $true
    Disconnect-EXO
}

function Remove-ExoMailboxPermission{
    [CmdLetBinding()]

	param( [string]$TargetMailboxOwner,
		   [String]$User
		   )
		Connect-EXO
    Remove-MailboxPermission -Identity $TargetMailboxOwner -User $User -AccessRights FullAccess -InheritanceType All -confirm:$False
    Disconnect-EXO

}

function Add-CSVtoO365group{
    [CmdLetBinding()]

    param(
    [Parameter(Mandatory=$True)]$FilePath,
	[Parameter(Mandatory=$True)]$GroupName
    )
    Connect-EXO
    Import-CSV $FilePath | 
    ForEach-Object{ Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $_.member }
    disconnect-exo
}

function Add-O365GroupUser{
    [CmdLetBinding()]
    param(
    [Parameter(Mandatory=$True)]$GroupName,
    [Parameter(Mandatory=$True)]$upn    
    )
    connect-EXO
    Add-UnifiedGroupLinks -Identity $GroupName -LinkType Members -Links $upn
    Write-host "Adding $upn to the group: $GroupName"
    Disconnect-EXO
}

function Get-DemoSqlConnectionString{
    [CmdLetBinding()]
    param()
    return "Data Source=tcp:rmxdemo.database.windows.net,1433;Initial Catalog=RMX-Demo;Authentication=Active Directory Integrated;"
  }
 
  function Update-Printers  {
      [cmdletbinding()]
      param(
          [Parameter(Mandatory=$False)]$ComputerName,
          [Parameter(Mandatory=$False)]$PrinterDepartment,
          [Parameter(Mandatory=$False)]$Color
          )
          $Printer = "*C2503*","*C5502*"

          if(!($ComputerName))
          {
              $ComputerName = read-host -prompt "Enter the Computer Name:"
          }
          if(!($PrinterDepartment))
          {
              $PrinterDepartment = read-host -prompt "Which printers are to be added? (Main, Billing, or Service):"
          }
          if(!($Color))
          {
              $Color = read-host -prompt "Is the user allowed to print in color? (Yes or No):"
          }
  
      Get-Printer -computername $ComputerName
      Remove-Printer -ComputerName $ComputerName -Name $printer
      
      Switch($PrinterDepartment)
      {
          'Main'
          {
              if($color -like 'N')
              {
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 BW\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.txt" -Destination \\$ComputerName\ADMIN$
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 BW\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe" -destination \\$ComputerName\ADMIN$
                  Invoke-Psexec -computername $ComputerName -Command "Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.txt"
              }
              else
              {
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 Color\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.txt" -Destination \\$ComputerName\ADMIN$
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 Color\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe" -destination \\$ComputerName\ADMIN$
                  Invoke-Psexec -computername $ComputerName -Command "Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.txt"
              }
          }
          'Billing'
          { 
              copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\Billing Color\Billing-_Aficio_MP_C5502-TCP_IP-RICOH_Aficio_MP_C5502_PCL_6-64Bit-for64bitOS-1.1.0.txt" -Destination \\$computer\ADMIN$
              copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\Billing Color\Billing-_Aficio_MP_C5502-TCP_IP-RICOH_Aficio_MP_C5502_PCL_6-64Bit-for64bitOS-1.1.0.exe" -destination \\$computer\ADMIN$
              Invoke-Psexec -computername $ComputerName -Command "Billing-_Aficio_MP_C5502-TCP_IP-RICOH_Aficio_MP_C5502_PCL_6-64Bit-for64bitOS-1.1.0.exe"
              remove-item "Billing-_Aficio_MP_C5502-TCP_IP-RICOH_Aficio_MP_C5502_PCL_6-64Bit-for64bitOS-1.1.0.txt"
              remove-item "Billing-_Aficio_MP_C5502-TCP_IP-RICOH_Aficio_MP_C5502_PCL_6-64Bit-for64bitOS-1.1.0.exe"
  
              if($color -like 'N')
              {
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 BW\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.txt" -Destination \\$ComputerName\ADMIN$
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 BW\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe" -destination \\$ComputerName\ADMIN$
                  Invoke-Psexec -computername $ComputerName -Command "Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.txt"
              }
              else
              {
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 Color\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.txt" -Destination \\$ComputerName\ADMIN$
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 Color\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe" -destination \\$ComputerName\ADMIN$
                  Invoke-Psexec -computername $ComputerName -Command "Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.txt"
              }
          }
          'Service'
          {
              Add-PrinterPort -ComputerName $ComputerName -Name "XeroxTCPPort:" -PrinterHostAddress "10.1.10.225"
              invoke-psexec -ComputerName $ComputerName -command "pnputil /add-driver '\\rocinante\shared\Printer Drivers\Xerox WorkCentre 6515\Xerox WorkCentre 6515\XeroxPhaser6510_WC6515_PCL6.inf'" 
              Add-PrinterDriver -ComputerName $ComputerName -Name "Xerox WorkCentre 6515 V4 PCL6" 
              Add-Printer -ComputerName $ComputerName -Name "Xerox Workstation 6515" -DriverName "Xerox WorkCentre 6515 V4 PCL6" -PortName "XeroxTCPPort:"
  
              if($color -like 'N')
              {
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 BW\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.txt" -Destination \\$ComputerName\ADMIN$
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 BW\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe" -destination \\$ComputerName\ADMIN$
                  Invoke-Psexec -computername $ComputerName -Command "Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.5.0.txt"
              }
              else
              {
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 Color\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.txt" -Destination \\$ComputerName\ADMIN$
                  copy-item "\\rocinante\shared\Printer Drivers\Ricoh Printer Drivers\Users_Package\C2503 Color\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe" -destination \\$ComputerName\ADMIN$
                  Invoke-Psexec -computername $ComputerName -Command "Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.exe"
                  remove-item "\\$ComputerName\admin$\Ricoh_C2503-TCP_IP-LANIER_MP_C2503_PCL_6-64Bit-for64bitOS-1.6.0.txt"
              }
          }
          'Canon'
          {
            Add-PrinterPort -ComputerName $ComputerName -Name "CanonTCPPort:" -PrinterHostAddress "10.2.10.222"
            invoke-psexec -ComputerName $ComputerName -command "pnputil /add-driver '\\rocinante\shared\Printer Drivers\Canon C5535\Generic Plus PCL6\etc\Cnp60MA64.inf'" 
            Add-PrinterDriver -ComputerName $ComputerName -Name "Canon Generic Plus PCL6" 
            Add-Printer -ComputerName $ComputerName -Name "Canon C5535 III" -DriverName "Canon Generic Plus PCL6" -PortName "CanonTCPPort:"
          }
      }
      
  }

function Update-Productkey{
$License = (Get-WmiObject -query ‘select * from SoftwareLicensingService’).OA3xOriginalProductKey
slmgr.vbs /ipk $license
start-sleep 30
slmgr.vbs /ato
}

function Add-AdminUser{
net user /add rmxadmin Lozhkin1!
WMIC USERACCOUNT WHERE "Name='rmxadmin'" SET PasswordExpires=FALSE
net localgroup administrators rmxadmin /add
net user administrator /active:no 
}

function Add-RMXVpn{
Add-VpnConnection -Name Rhythmedix -ServerAddress 50.228.161.254 -AllUserConnection $true -SplitTunneling $false -authenticationmethod mschapv2 -tunneltype l2tp -l2tppsk 73EE1DB45064CFF9 -encryptionlevel Required -passthru
}

function ProvisionBitlocker{
    Manage-BDE -On C: -SkipHardwareTest -ComputerName $env:COMPUTERNAME
    $RecoveryKey = Get-BitLockerVolume -MountPoint C: | Select-Object -ExpandProperty KeyProtector | Where-Object KeyProtectorType -eq 'RecoveryPassword'
# In case there is no Recovery Password, lets create new one
if (!$RecoveryKey)
	{
	Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
	$RecoveryKey = Get-BitLockerVolume -MountPoint C: | Select-Object -ExpandProperty KeyProtector | Where-Object KeyProtectorType -eq 'RecoveryPassword'
	}
    Out-File -InputObject $RecoveryKey -FilePath '\\rocinante\shared\KeyBackupTempFolder\Key.txt'
}

function Add-WVDAppUser{
    param(
    [parameter(Mandatory=$True)]$user
    )
    $upn = $user + "@rhythmedix.com"  

    Add-AdGroupMember -Identity "Azure AD Domain Services" -Members (Get-ADUser -filter {EmailAddress -eq $upn})
    Add-AdGroupMember -Identity "Remote Portal Users" -Members (Get-ADUser -filter {EmailAddress -eq $upn})
    Sync-Azure
       
    Add-WVDIps -upn $upn


}

function Add-WVDIps
{
   param(
       [Parameter(Mandatory=$True)]$upn
   ) 
    $ip = '52.224.14.217'
    Add-IPWhitelist -UPN $upn -IP $ip
    $ip = '52.224.14.214'
    Add-IPWhitelist -UPN $upn -IP $ip
    $ip = '52.224.14.199'
    Add-IPWhitelist -UPN $upn -IP $ip
}


function Invoke-WVDUserDisconnect{
    param(
        [parameter(Mandatory=$True)]$user
        )
        $upn = $user + "@rhythmedix.com"    
    Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
    Get-RdsUserSession -TenantName "RhythMedix Remote Review" -HostPoolName "RemoteReview_HostPool" | where-object { $_.UserPrincipalName -eq $upn } | Invoke-RdsUserSessionLogoff -NoUserPrompt
}

function Get-WVDUsers{
    $rdsappgroup = "Remote Review"
    $hostpool = "RemoteReview_HostPool"
    $tenantname = "RhythMedix Remote Review"
    Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
    Get-RdsAppGroupUser -TenantName $tenantname -HostPoolName $hostpool -AppGroupName $rdsappgroup | Out-Host
}

function Add-WVDDestkopUser{
    param(
    [parameter(Mandatory=$True)]$UPN
    )
    $rdsappgroup = "Desktop Application Group"
    $hostpool = "RemoteReview_HostPool"
    $tenantname = "RhythMedix Remote Review"
    #Add-AdGroupMember -Identity "Azure AD Domain Services" -Members (Get-ADUser -filter {EmailAddress -eq $upn})
    #Sync-Azure
    Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
    Remove-RdsAppGroupUser -TenantName $tenantname -HostPoolName $hostpool -AppGroupName "Remote Review" -UserPrincipalName $upn
    Add-RdsAppGroupUser -TenantName $tenantname -HostPoolName $hostpool -AppGroupName $rdsappgroup -UserPrincipalName $upn
  
}

function New-WVDRemoteApp{
    param(
    [parameter(Mandatory=$true)]$rdsappgroup,  #"AppGroupName"
    [parameter(Mandatory=$true)]$hostpool, #"RemoteReview_HostPool"
    [parameter(Mandatory=$true)]$tenantname, # "RhythMedix Remote Review"
    [parameter(Mandatory=$true)]$filepath, # "C:\File\Path.extension"
    [parameter(Mandatory=$true)]$iconpath, # "C:\Windows\system32\mstsc.exe"
    [parameter(Mandatory=$true)]$rdsappname, # "AppName"
    [parameter(Mandatory=$true)]$rdsappfriendlyname
     )# "App Name in List"
   #>
    Add-RdsAccount -DeploymentUrl "https://rdbroker.wvd.microsoft.com" -credential (get-storedcredential -target O365Admin)
   New-RDSAppGroup -TenantName $tenantname -HostPoolName $hostpool -AppGroupName $rdsAppGroup -ResourceType "RemoteApp"
   New-RDSRemoteApp -TenantName $tenantname -HostPoolName $Hostpool -AppGroupName $rdsAppGroup -Name $rdsappname -FilePath $filepath -FriendlyName $rdsappfriendlyname -IconPath $iconpath
   <#
    $rdsappgroup = "Holter Manager"
    New-RDSRemoteApp -TenantName $tenantname -HostPoolName $Hostpool -AppGroupName $rdsappgroup -Name $rdsappname -FilePath $filepath -FriendlyName $rdsappfriendlyname -IconPath $iconpath
    
  Add-AdGroupMember -Identity "Azure AD Domain Services" -Members (Get-ADUser -filter {EmailAddress -eq $upn})
  
    #Remove-RdsAppGroupUser -TenantName $tenantname -HostPoolName $hostpool -AppGroupName "Desktop Application Group" -UserPrincipalName $upn
    #Add-RdsAppGroupUser -TenantName $tenantname -HostPoolName $hostpool -AppGroupName $rdsappgroup -UserPrincipalName $upn
#>   
}

function Disable-User{
    [CmdletBinding()]
    param (
        
        [Parameter()]$User
    )
    #-DateTime 'mm:dd:yyyy hh:mm:ss'
    [DateTime]$WhenToDisable = Read-Host "What day and time should the user be disabled? (format 'mm/dd/yyyy hh:mm:ss')" 
    Invoke-Command -ComputerName galactica.rhythmedix.com -ScriptBlock{
        param($user,$WhenToDisable)
        $tasktrigger = New-ScheduledTaskTrigger -Once -at $WhenToDisable
        $taskprincipal =  New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest -LogonType ServiceAccount
        $taskaction = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument "-ExecutionPolicy ByPass -File C:\Disable-ADUser.ps1 $user"
        
        Register-ScheduledTask -TaskName "Disable $user" -Trigger $tasktrigger -Action $taskaction -Principal $taskprincipal
    } -ArgumentList $user,$WhenToDisable
}

function Set-AzureComputerSync{
    get-adcomputer $env:computername | Add-ADGroupMember "Azure AD Sync"
}
 
function Update-AssociateIDs{
    [CmdletBinding()]
    param (
        
        [Parameter(Mandatory=$false)]$CSVFile,
        [Parameter(Mandatory=$false)]$SamAccountName,
        [Parameter(Mandatory=$false)]$AssociateID
        
    )

    if($CSVFile)
    {
        $users = import-csv $CSVFile
        ForEach($user in $users)
        {
            $username = $user.FirstName + " " + $user.LastName
            $username
            try 
            {
                Get-ADuser -filter {Displayname -like $username} | set-aduser -Replace @{
                    'AssociateId' = $user.AssociateID;
                }
            }
            catch {
                Out-Host $Username + " update failed"
            }
        
        }
    }
    else {
        get-ADuser $SamAccountName | set-aduser -Replace @{
            'AssociateId' = $AssociateID;
        }
    }
}

function Convert-CloudUserToADSYnc{
[CmdletBinding()]
param (
    [Parameter(Mandatory=$true)]
    [string]    
    $upn,
    [Parameter(Mandatory=$true)]
    [string]    
    $title,
    [Parameter(Mandatory=$true)]
    [string]    
    $manager,
    [Parameter(Mandatory=$true)]
    [string]    
    $department
)   
    Connect-MsolService -credential (get-storedcredential -target O365Admin)
    $clouduser = get-msoluser -userprincipalname $upn
    $givenname = $clouduser.FirstName
    $surname = $clouduser.LastName
    $pattern = '[^a-zA-Z]'
    $samaccountname = ($givenname[0] + ($surname -replace $pattern, '') + $suffix).tolower()
    $tempPassword = convertto-securestring "Password1" -asplaintext -force
    $displayname = $givenname + " $surname"

    #$newTempPassword = convertto-securestring "Rhythmedix1!" -asplaintext -force

    New-AdUser -Name $displayName -SamAccountName $samaccountName -AccountPassword $tempPassword -ChangePasswordAtLogon $false -Department $department -Title $title -DisplayName $displayName -EmailAddress $upn -GivenName $givenName -Surname $surname -UserPrincipalName $upn -EmployeeID $empNumber -Enabled $true -PassThru
    $ID= [system.convert]::ToBase64String((Get-ADUser -filter {userprincipalname -eq $UPN}).objectGUid.ToByteArray())
    Set-MsolUser -UserPrincipalName $upn -ImmutableId $ID
    Add-ADGroupMember -identity 'Azure AD Sync'
    Sync-Azure
}

Function Set-WebGLStatus{
[CmdLetBinding()]
    Param(
    [Parameter(Mandatory=$True)]$UPN,
    [Parameter(Mandatory=$true)]$ENABLED
    )
    if($ENABLED = $true)
    {
        $flag = 1
    }
    else {
        $flag = 0
    }

    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = Get-SqlConnectionString
        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.CommandText = "dbo.spSystemSetWebGlFlag"  ## this is the stored proc name 
        $SqlCmd.Connection = $SqlConnection  
        $SqlCmd.CommandType = [System.Data.CommandType]::StoredProcedure  ## enum that specifies we are calling a SPROC
        #SP format exec [spSystemSetWebGlFlag] @UserName = 'vzobrak@rhythmedix.com', @enabled = 1
        $param1=$SqlCmd.Parameters.Add("@USERNAME" , [System.Data.SqlDbType]::VarChar)
            $param1.Value = $UPN 
        $param2=$SqlCmd.Parameters.Add("@ENABLED" , [System.Data.SqlDbType]::VarChar)
            $param2.Value = $flag
        $SqlConnection.Open()
        $result = $SqlCmd.ExecuteNonQuery() 
        Write-output "result=$result" 
        $SqlConnection.Close()
    
}

function New-ShortUrl{
    param(
        [Parameter(Mandatory=$true)]$longurl,
        [ValidateLength(8,[int]::MaxValue)][Parameter(Mandatory=$false)]$alias
    )
if($alias -eq "")
{
    invoke-WebRequest -UseBasicParsing https://rmx.health/ShortenUrl?code=7rxIKBuRNQ8fQBPvHgdH52jYgdupvmHLbLRrTtheyzJ9/Wr/1LaKvw== -ContentType "application/json" -Method POST -Body "{ 'longUrl':`'$longurl`'}"
}
else {
    invoke-WebRequest -UseBasicParsing https://rmx.health/ShortenUrl?code=7rxIKBuRNQ8fQBPvHgdH52jYgdupvmHLbLRrTtheyzJ9/Wr/1LaKvw== -ContentType "application/json" -Method POST -Body "{ 'longUrl':`'$longurl`', 'alias':`'$alias`'}"
}
}