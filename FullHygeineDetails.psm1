#User Logging Credentials
  $trgtOrgUrl = "https://technoyash.crm8.dynamics.com"
  $trgtUserName = "yash@technoyash.onmicrosoft.com"
  $trgtPass = ConvertTo-SecureString "qwerty@1234" -AsPlainText -Force
  $instaceid = "51976e13-5a20-4a55-9f47-70924943d136"
  $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $trgtUserName, $trgtPass
  $CurrentDate = Get-Date
  $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh:mm:ss')
  $CurrentDate1 = Get-Date
  $CurrentDate1 = $CurrentDate1.ToString('MM-dd-yyyy')
  
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
#$scriptPath

#SendEmail is used to send multiple attachments to the User
function SendEmail{
param(
[string] $loc1,
[string] $loc2,
[string] $loc3
)
try{
 #Credentials, Subject, Body and SMTP Details
   $EmailFrom = "yash.gupta@soprasteria.com"
   $EmailTo = "yash.gupta@soprasteria.com"
   $SMTPServer = "ptx.send.corp.sopra" 
   $Subject = $CurrentDate1 + " Hygeine Data"
   $Body = "Hello User,
Please find Attachment of "+$CurrentDate1 +" Hygeine Data.
          
Thanks and Regards,
Administration"
   $SMTPAuthUsername = "" #write your Emailid
   $SMTPAuthPassword = "" #Write your Password
   $mailmessage = New-Object system.net.mail.mailmessage 
   $mailmessage.from = ($emailfrom) 
   $mailmessage.To.add($emailto)
   $mailmessage.Subject = $Subject
   $mailmessage.Body = $Body
   $attachment = New-Object System.Net.Mail.Attachment($loc1)
   $mailmessage.Attachments.Add($attachment)
   $attachment = New-Object System.Net.Mail.Attachment($loc2)
   $mailmessage.Attachments.Add($attachment)
   $attachment = New-Object System.Net.Mail.Attachment($loc3)
   $mailmessage.Attachments.Add($attachment)
   #$mailmessage.IsBodyHTML = $true
   $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25)  
   $SMTPClient.Credentials = New-Object System.Net.NetworkCredential("$SMTPAuthUsername", "$SMTPAuthPassword") 

   $SMTPClient.Send($mailmessage)
  $CurrentDate + " Email Send Successfully to " +$EmailTo | Add-Content '.\HygeineLogging.log'
  }
   catch
   {
 #  [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','SendEmail','Ok','Error')
 
      if(Test-Path '.\HygeineLogging.log'){
        $CurrentDate | Add-Content '.\HygeineLogging.log'
        $_.Exception.Message | Out-File '.\HygeineLogging.log'
      
    }
    else{
        $CurrentDate | Add-Content '.\HygeineLogging.log'
        $_.Exception.Message | Out-File '.\HygeineLogging.log'
        
            }
    exit
   }
}





# "GetBackUpDetails" is used to export the backup history of the live instance

function GetBackUpDetails{
try{

  #$trgtCRMOrg = Connect-CrmOnline -Credential $cred -ServerUrl $trgtOrgUrl

    $CurrentDate + " User Logged in Successfully in " + $trgtOrgUrl | Add-Content '.\HygeineLogging.log'
    
 #Region Url
    $Api = "https://admin.services.crm8.dynamics.com"

 #Getting BackUp Details
    $backup=Get-CrmInstanceBackups -ApiUrl $Api -Credential $cred -InstanceId $instaceid
    

    $result=$backup | Select-Object -Property Date,Id,Status,CreatedBy,CreatedOn,ExpiresOn,Version,Notes,Label
   
   $CurrentDate + " BackUp retreived Successfully of " +$trgtOrgUrl | Add-Content '.\HygeineLogging.log'

   #Creating Tablubar format to store in CSV File
    $tableEntity = New-Object system.Data.DataTable "AsyncJobs"
    $tblcol1 = New-Object system.Data.DataColumn Extract_Date, ([string])
    $tblcol2 = New-Object system.Data.DataColumn Id, ([string])
    $tblcol3 = New-Object system.Data.DataColumn Status, ([string])
    $tblcol4 = New-Object system.Data.DataColumn CreatedBy, ([string])
    $tblcol5 = New-Object system.Data.DataColumn CreatedOn, ([string])
    $tblcol6 = New-Object system.Data.DataColumn ExpiresOn, ([string])
    $tblcol7 = New-Object system.Data.DataColumn Version, ([string])
    $tblcol8 = New-Object system.Data.DataColumn Notes, ([string])
    $tblcol9 = New-Object system.Data.DataColumn Label, ([string])
    $tableEntity.columns.add($tblcol1)
    $tableEntity.columns.add($tblcol2)
    $tableEntity.columns.add($tblcol3)
    $tableEntity.columns.add($tblcol4)
    $tableEntity.columns.add($tblcol5)
    $tableEntity.columns.add($tblcol6)
    $tableEntity.columns.add($tblcol7)
    $tableEntity.columns.add($tblcol8)
    $tableEntity.columns.add($tblcol9)

  #Inserting data into the Table
        $result | ForEach-Object {
     $tblrow = $tableEntity.NewRow()
          $tblrow.Extract_Date = $CurrentDate
          $tblrow.Id = $_.Id
          $tblrow.Status = $_.Status
          $tblrow.CreatedBy = $_.CreatedBy
          $tblrow.CreatedOn = $_.CreatedOn
          $tblrow.ExpiresOn = $_.ExpiresOn
          $tblrow.Version = $_.Version
          $tblrow.Notes = $_.Notes
          $tblrow.Label = $_.Label

    $tableEntity.Rows.Add($tblrow)        
    }
 
  #Adding the backUp details to Csv file   
   $tableEntity | Export-Csv '.\BackupDetails.csv' -Append

   $CurrentDate + " The BackUp Data has been exported."  | Add-Content '.\HygeineLogging.log'

 #  [System.Windows.MessageBox]::Show('BackUp data Retreived Successfull!! Please See Target File for information','Backup Details','Ok','Information')
   
  $CurrentDate + " GetBackUpDeatails Module Run Successfully" | Add-Content '.\HygeineLogging.log'
    }

 #catching exceptions
    catch [System.Net.WebException]
  {
     $CurrentDate | Add-Content '.\HygeineLogging.log'
     $_.System.Net.WebException | Set-Content '.\HygeineLogging.log'
     [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{        
    $CurrentDate | Add-Content '.\HygeineLogging.log'
    "Could not find path" | Add-Content '.\HygeineLogging.log'
    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
      $CurrentDate | Add-Content '.\HygeineLogging.log'
      $_.System.IO.IOException | Add-Content '.\HygeineLogging.log'
   #   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
  }
    catch
   {
   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
 
      if(Test-Path '.\HygeineLogging.log'){
        $CurrentDate | Add-Content '.\HygeineLogging.log'
        $_.Exception.Message | Out-File '.\HygeineLogging.log'
      
    }
    else{
        $CurrentDate | Add-Content '.\HygeineLogging.log'
        $_.Exception.Message | Out-File '.\HygeineLogging.log'
        
            }
    exit
   }

}
GetBackUpDetails






# This Function is used get the details of licenses
function LicenceDetails{
try{
 #Region Url
    $Api = "https://admin.services.crm8.dynamics.com"
 #Connecting O365
    Connect-MsolService -Credential $cred


 #Getting License Details
    $result=  Get-MsolAccountSku | Select-Object -Property Date,AccountSkuId,ActiveUnits,ConsumedUnits,WarningUnits

    $CurrentDate + " License details retreived Successfully" | Add-Content '.\HygeineLogging.log'

 #Creating Tablubar format to store in CSV File
    $tableEntity = New-Object system.Data.DataTable "AsyncJobs"
    $tblcol1 = New-Object system.Data.DataColumn Extract_Date, ([string])
    $tblcol2 = New-Object system.Data.DataColumn AccountSkuId, ([string])
    $tblcol3 = New-Object system.Data.DataColumn ActiveUnits, ([string])
    $tblcol4 = New-Object system.Data.DataColumn ConsumedUnits, ([string])
    $tblcol5 = New-Object system.Data.DataColumn WarningUnits, ([string])
    $tableEntity.columns.add($tblcol1)
    $tableEntity.columns.add($tblcol2)
    $tableEntity.columns.add($tblcol3)
    $tableEntity.columns.add($tblcol4)
    $tableEntity.columns.add($tblcol5)

  #Inserting Data into the Table
        $result | ForEach-Object {
     $tblrow = $tableEntity.NewRow()
          $tblrow.Extract_Date = $CurrentDate
          $tblrow.AccountSkuId = $_.AccountSkuId
          $tblrow.ActiveUnits = $_.ActiveUnits
          $tblrow.ConsumedUnits = $_.ConsumedUnits
          $tblrow.WarningUnits = $_.WarningUnits

          $tableEntity.Rows.Add($tblrow)        
    }
    
  #Getting license Details and export into csv file
    $tableEntity | Export-Csv '.\LicenseDetails.csv' -Append
    $CurrentDate + " The License Data has been exported." | Add-Content '.\HygeineLogging.log'

  #  [System.Windows.MessageBox]::Show('Operation Successfull!! Please See Target File for information','Backup Details','Ok','Information')

    $CurrentDate + " LicenseDetails Module run Successfully Successfully" | Add-Content '.\HygeineLogging.log'
}

 #catching exceptions
  catch [System.Net.WebException]
  {
   
     $CurrentDate | Add-Content '.\HygeineLogging.log'
     $_.System.Net.WebException | Set-Content '.\HygeineLogging.log'
   #  [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{        
    $CurrentDate | Add-Content '.\HygeineLogging.log'
    "Could not find path" | Add-Content '.\HygeineLogging.log'
    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
      $CurrentDate | Add-Content '.\HygeineLogging.log'
      $_.System.IO.IOException | Add-Content '.\HygeineLogging.txt'
      [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
  }
    catch
   {
     $CurrentDate | Add-Content '.\HygeineLogging.log'
     [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
      if(Test-Path '.\HygeineLogging.log'){
        Add-Content '.\HygeineLogging.log' -Value $CurrentDate
        $_.Exception.Message | Out-File '.\HygeineLogging.log'
        
    }
    else{
        $CurrentDate | Add-Content '.\HygeineLogging.log' 
        $_.Exception.Message | Out-File '.\HygeineLogging.log'
        
            }
    exit
   }
}

LicenceDetails






#This Function is getting the details of the waiting and failed System Job
function getsystemjob{

try{
 #User Logging Credentials
    $trgtOrgUrl = "https://technoyash.crm8.dynamics.com"
    $trgtUserName = "yash@technoyash.onmicrosoft.com"
    $trgtPass = ConvertTo-SecureString "qwerty@1234" -AsPlainText -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $trgtUserName, $trgtPass
  #$orgName="Yash Technologies"
  #$con = Get-CrmConnection -OrganizationName $orgName -OnLineType Office365 -Credential $cred

 #Connecting CRM
   $trgtCRMOrg = Connect-CrmOnline -Credential $cred -ServerUrl $trgtOrgUrl

 #Fetch XML on Entity System Job
   $fetchXml = @"


<fetch>
  <entity name="asyncoperation" >
    <attribute name="name" />
    <attribute name="statuscode" />
    <attribute name="createdon" />
    <attribute name="friendlymessage" />
    <filter type="and" >
      <filter type="or" >
        <condition attribute="statuscode" operator="eq" value="10" />
        <condition attribute="statuscode" operator="eq" value="31" />
      </filter>
    </filter>
  </entity>
</fetch>

"@
  
  $CurrentDate + " Waiting and failed System Jobs retreived Successfully of " +$trgtOrgUrl | Add-Content '.\HygeineLogging.log'

 #Creating Tablubar format to store in CSV File
    $tableEntity = New-Object system.Data.DataTable "AsyncJobs"
    $tblcol1 = New-Object system.Data.DataColumn Extract_Date, ([string])
    $tblcol2 = New-Object system.Data.DataColumn name, ([string])
    $tblcol3 = New-Object system.Data.DataColumn statuscode, ([string])
    $tblcol4 = New-Object system.Data.DataColumn friendlymessage, ([string])
    $tableEntity.columns.add($tblcol1)
    $tableEntity.columns.add($tblcol2)
    $tableEntity.columns.add($tblcol3)
    $tableEntity.columns.add($tblcol4)
    $FetchResult = Get-CrmRecordsByFetch -conn $trgtCRMOrg -Fetch $fetchXml -AllRows
   # $solutionEntity = New-Object System.Collections.Generic.List[Guid]

    $FetchResult.CrmRecords | ForEach-Object {
     $tblrow = $tableEntity.NewRow()
          $tblrow.Extract_Date = $CurrentDate
          $tblrow.name = $_.name
          $tblrow.statuscode = $_.statuscode
          $tblrow.friendlymessage = $_.friendlymessage
          
          $tableEntity.Rows.Add($tblrow)        
    }
    
 #Exporting Data in CSV File
    
    $tableEntity | Export-Csv -Path '.\systemJob.csv' -Append

    $CurrentDate + " The System Job Data has been exported." | Add-Content '.\HygeineLogging.log'
  #  [System.Windows.MessageBox]::Show('Operation Successfull!! Please See Target File for information','Backup Details','Ok','Information')

   $locSystem = ".\systemJob.csv"
   $locBack = ".\BackupDetails.csv"
   $locLicense = ".\LicenseDetails.csv"


 #callibg SendEmail function to send Email to the user
   $CurrentDate + " getsystemjob Module run successfully" | Add-Content '.\HygeineLogging.log'



  }

 #catching exceptions
  catch [System.Net.WebException]
  {
     $CurrentDate | Add-Content '.\HygeineLogging.log' 
     $_.System.Net.WebException | Set-Content '.\HygeineLogging.log'
 #    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{      
    $CurrentDate | Add-Content '.\HygeineLogging.log'   
    $_.System.IO.FileNotFoundException | Add-Content '.\HygeineLogging.log'
    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
      $CurrentDate | Add-Content '.\HygeineLogging.log' 
      $_.System.IO.IOException | Add-Content '.\HygeineLogging.log'
   #   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
  }
    catch
   {
   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
  
      if(Test-Path '.\HygeineLogging.log'){
        Add-Content '.\getsystemjobErrorLog1.log' -Value $CurrentDate
        $_.Exception.Message | Out-File '.\HygeineLogging.log'      
    }
    else{
        Add-Content -Path '.\HygeineLogging.log' -Value $CurrentDate
        $_.Exception.Message | Out-File '.\getsystemjobErrorLog.log'
        
            }
    exit
   }
  }

 getsystemjob



SendEmail $locBack $locLicense $locSystem
