Install-Module -Name InvokePsExec
Install-Module -Name CredentialManager

function Get-SqlConnectionString(){
    return "Data Source=tcp:rmxprod.database.windows.net,1433;Initial Catalog=rhythmedix;Authentication=Active Directory Integrated;"
  }
  
  Function Add-RhythmstarUser{
      Param(
      [Parameter(Mandatory=$True)]$FullName,
      [Parameter(Mandatory=$false)]$Demo=$false
      )
      if($demo -eq $False)
      {
          $proceed = read-host -prompt "Add: $FullName to the Clinical Rhythmstar Portal? Y/N"
      }
      else
       {
          $proceed = read-host -prompt "Add: $FullName to the Demo Rhythmstar Portal? Y/N"
      }
      if($proceed -like 'Y')
      {
        $FirstInitial = $FullName.Substring(0,1)
        $FirstName, $LastName = $FullName -split "\s", 2
        $accountName = $FirstInitial + $LastName #login name
        $UPN = $accountName.ToLower() + "@rhythmedix.com" #userprincipalname
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        if($Demo -eq $True)
        {
            $SqlConnection.ConnectionString = Get-DemoSqlConnectionString
        }
        else 
        {
            $SqlConnection.ConnectionString = Get-SqlConnectionString
        }   
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
  
  
  function Add-NewUser{
      param(
      [Parameter(Mandatory=$True)]$FullName,
      [Parameter(Mandatory=$True)]$Title,
      [Parameter(Mandatory=$False)]$Department,
      [Parameter(Mandatory=$False)]$Manager,
      [Parameter(Mandatory=$False)]$Ladies,
      [Parameter(Mandatory=$False)]$Rhythmstar,
      [Parameter(Mandatory=$False)]$Location
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
              AD P1 - holter, sales, VPs and anyone else who has a laptop/VPN
  
          To order/change license count - send email to Strong, Katrina <Katrina.Strong@softwareone.com> with number of licenses to add/remove. 
          We have our enterprise agreement with Microsoft through software one.
      #>
  
      $FirstInitial = $FullName.Substring(0,1)
      $FirstName, $LastName = $FullName -split "\s", 2
      $accountName = $FirstInitial + $LastName #login name
      $accountName = $accountName.Tolower()
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
          'Sales'
              {	
                  $Title = 'Sales Representative'
                  $Department = 'Sales'
                  $Manager = 'kgartland'
                  break
              }
      }
                  #'Clinical'
      #"Clinical" #Clinical, Logistics, IT, Sales, Payer Relations, Clinical Administrators, Engineering and Development
      #$title = "Arrhythmia Analyst" #"Product Distribution Specialist" #"Clinical Administrator" #"Arrhythmia Analyst" #"Holter Technician" #"Sr. Arrhythmia Analyst"
      #$manager = "tcatling" #don't need UPN here, just login name (first portion) #tcatling, evalentine, arichmann, ndemiranda
  
      $displayName = $firstName + " " + $lastName
      $upn = $accountName + "@rhythmedix.com" #userprincipalname
      $email = $upn
      $tempPassword = convertto-securestring "Password1" -asplaintext -force
      if(!($Department))
      {
          $Department = read-host -prompt "What department is $fullname in:"
      }
      if(!($Manager))
      {
          $Manager= read-host -prompt "Who is $fullname's manager:"
      }
      Write-Host "
                  User Name: $displayname
                  Title: $title
                  Department: $Department
                  Manager: $Manager
                  Email: $email "
        $continue = read-host -Prompt "Continue? Y/N"
    }while($continue -like 'N')
      $user = New-AdUser -Name $displayName -SamAccountName $accountName -AccountPassword $tempPassword -ChangePasswordAtLogon $true -Department $department -Title $title -DisplayName $displayName -EmailAddress $email -GivenName $firstName -Surname $lastName -Manager $manager -UserPrincipalName $upn -Enabled $true -PassThru
  
      #common for everyone
      #Add-AdGroupMember "All Employees" $user
      Add-AdGroupMember "Azure AD Sync" $user #required group to sync to cloud
      
      Connect-MsolService -Credential $credential
  
      DO
      {		
              Sync-Azure
              Write-Host "." -NoNewline
              Start-Sleep -Seconds 10
      } Until (Get-MsolUser -UserPrincipalName $UPN -ErrorAction SilentlyContinue)
      
      if($location -ne "Leominster")
      {
          Add-AdGroupMember "Mount Laurel Office" $user
      }
      else
      {
          Add-AdGroupMember "Tolman Office" $user #if user is in MA (not hiring there right now)
      }
      
      Set-MsolUser -UserPrincipalName $upn -UsageLocation "US"
      Add-O365GroupUser -GroupName "All Employees" -Username $upn
      
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
                  Add-AdGroupMember "VPN Users" $user
                  Add-AdGroupMember "Self-Service Password Reset" $user
                  Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:AAD_Premium"
                              
              }
              else
              {
                  #monitoring center - arrhythmia analyst + sr. arrhythmia analyst
                  Add-AdGroupMember "Hourly Employees" $user
                #  Add-AdGroupMember "Monitoring" $user
                  Add-O365GroupUser -GroupName "RMX Monitoring" -Username $upn
                  Add-RhythmstarUser -FullName $FullName
                  if($location -eq "Leominster")
                  {
                      add-o365groupuser -GroupName "Leominster Monitoring" -Username $upn
                  }
                  else
                  {
                      Add-O365GroupUser -GroupName "Mt Laurel Monitoring" -Username $upn
                  }
              }
          }
          'Clinical Administrators'
          {
              #clinical admin
                  Add-AdGroupMember "Hourly Employees" $user
                  Add-AdGroupMember "Ringcentral Softphone Users" $user
                  Add-RhythmstarUser -FullName $FullName
          }
          'Logistics'
          {
              #logistics
                  Add-AdGroupMember "Hourly Employees" $user
                  Add-RhythmstarUser -FullName $FullName
                  
          }
          'Sales'
          {
              #sales
                  Add-AdGroupMember "Regulatory Medical Device" $user
                  Add-AdGroupMember "VPN Users" $user
                  Add-AdGroupMember "Self-Service Password Reset" $user
                  Add-RhythmstarUser -FullName $FullName -Demo:$True
                  Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:AAD_Premium"
          }
          'Engineering'
          {
              #engineering
                  Add-AdGroupMember "Regulatory Medical Device" $user
                  Add-AdGroupMember "RPSS User" $user
          }
          'IT'
          {
              #IT
                  Add-O365GroupUser -GroupName "Rhythmedix IT" -Username $upn
                  Add-AdGroupMember "VPN Users" $user
                  Add-AdGroupMember "Self-Service Password Reset" $user
                  Add-RhythmstarUser -FullName $FullName
                  Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:AAD_Premium"
                 
                  If($Title -eq "Web Developer")
                  {
                     Add-O365GroupUser -GroupName "Development" -Username $upn
                     Add-O365GroupUser -GroupName "Dev Team" -Username $upn
                  }
          }
  
      }
  
      
  
  
      if ($Department -eq "Logistics")
      {
          Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:StandardPACK"
      }
      else
      {
          Set-MsolUserLicense -UserPrincipalName $upn -AddLicenses "rhythmedix:ENTERPRISEPACK"
      }
  
      
  }
  function convert-ADUserToCloudOnly 
  {
      param
      (
      [Parameter(mandatory=$true)]$accountname
      )
      $samaccountname = get-aduser -Identity $accountname
      $upn = $SAMaccountname.userprincipalname
      Remove-ADGroupMember -Identity "Azure AD Sync" -Members $samaccountname
  
      Sync-Azure
      Connect-MsolService -Credential $credential
      Get-MsolUser -UserPrincipalName $upn -ReturnDeletedUsers | Restore-MsolUser
      Get-MsolUser -UserPrincipalName $upn | Set-MsolUser -ImmutableId ""
      
      Get-Aduser -Identity $accountname | Remove-ADUser
  }
  
  
  
  
  
  
  
  function Export-DLtoCSV{
   param(
      [Parameter(Mandatory=$True)]$GroupName
      )
      $DGName = $GroupName
          Get-DistributionGroupMember -Identity $DGName | Select-Object Name, PrimarySMTPAddress |
          Export-CSV "C:\\Distribution-List-Members.csv" -NoTypeInformation -Encoding UTF8
  }
  
  

  Function Sync-Azure{
      #PSFile version gives the on-screen feedback but requires ps1 file in the correct folder. The Command version doesn't give feedback but works regardless 
    Invoke-PsExec -ComputerName Galactica -PSFile "C:\powershell tools\Sync_Azure.ps1" 
    <#
    Invoke-PsExec -ComputerName Galactica -Command {
      Write-Host "Initializing Azure AD Delta Sync..." -ForegroundColor Yellow
  
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
    #>

  }

  function Connect-O365Compliance{
    if (!(get-pssession | where-object {$_.ConnectionURI -eq 'https://ps.compliance.protection.outlook.com/powershell-liveid/'}))
	{
		$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection
        Import-PSSession $Session -DisableNameChecking
    }
}

function Connect-EXO{
    #$UserCredential = Get-StoredCredential -Target O365Admin
    if (!(get-pssession | where-object {$_.ConfigurationName -eq 'Microsoft.Exchange'}))
	{
        $ExoSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $credential -Authentication Basic -AllowRedirection
        Import-PSSession $ExoSession -DisableNameChecking
	}

}

function Remove-Phishing{
	param(
		[Parameter(Mandatory=$True)]$SearchName
	)
    Connect-O365Compliance
	New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType HardDelete
}

function Disconnect-EXO{
	$ExoSession = get-pssession | where-object {$_.ConfigurationName -eq 'Microsoft.Exchange'}
	Remove-PSSession $ExoSession
}

function Disconnect-O365Compliance{
	$ComplianceSession = get-pssession | where-object {$_.ConfigurationURI -eq 'https://ps.compliance.protection.outlook.com/powershell-liveid/'}
	Remove-PSSession $ComplianceSession
}

function Add-ExoMailboxPermission{
	param( 
		[string]$MailboxOwner,
		[string]$User
		)
	connect-exo
    Add-MailboxPermission -Identity $MailboxOwner -User $User -AccessRights FullAccess -InheritanceType All -AutoMapping $true
}

function Remove-ExoMailboxPermission{

	param( [string]$TargetMailboxOwner,
		   [String]$User
		   )
		Connect-EXO
    Remove-MailboxPermission -Identity $TargetMailboxOwner -User $User -AccessRights FullAccess -InheritanceType All -confirm:$False

}

function Add-CSVtoO365group{

    param(
    [Parameter(Mandatory=$True)]$FilePath,
	[Parameter(Mandatory=$True)]$GroupName
    )
    Connect-EXO
    Import-CSV $FilePath | 
    ForEach-Object{ Add-UnifiedGroupLinks –Identity $GroupName –LinkType Members –Links $_.member }
}



function Add-O365GroupUser{
    param(
    [Parameter(Mandatory=$True)]$GroupName,
    [Parameter(Mandatory=$True)]$upn    
    )
    connect-EXO
    Add-UnifiedGroupLinks –Identity $GroupName –LinkType Members –Links $upn
    Write-host "Adding $upn to the group: $GroupName"
}


function Get-DemoSqlConnectionString{
    return "Data Source=tcp:rmxdemo.database.windows.net,1433;Initial Catalog=RMX-Demo;Authentication=Active Directory Integrated;"
  }
 
