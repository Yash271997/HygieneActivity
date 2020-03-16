
#SendEmail is used to send multiple attachments to the User
function Send-Email{
param(
[parameter(Position=1, Mandatory=$true)]
[string] $SystemJobloc,
[parameter(Position=2, Mandatory=$true)]
[string] $BackUploc,
[parameter(Position=3, Mandatory=$true)]
[string] $Licenseloc
)
try{
  $CurrentDate = Get-Date
  $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh:mm:ss')
  $CurrentDate1 = $CurrentDate.ToString('MM-dd-yyyy')
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
   $attachment = New-Object System.Net.Mail.Attachment($SystemJobloc)
   $mailmessage.Attachments.Add($attachment)
   $attachment = New-Object System.Net.Mail.Attachment($BackUploc)
   $mailmessage.Attachments.Add($attachment)
   $attachment = New-Object System.Net.Mail.Attachment($Licenseloc)
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
 
        $CurrentDate1 + "An Error has occured" | Add-Content '.\HygeineLogging.log'
        $_.Exception.Message | Out-File '.\Error.log'

    exit
   }
}



# "GetBackUpDetails" is used to export the backup history of the live instance

function Get-BackUpDetails{
param(
[parameter(Position=1, Mandatory=$true)]
[PSCredential]$Credential, 
[parameter(Position=2, Mandatory=$true)]
[string] $TargetOrganizationURL,
[parameter(Position=3, Mandatory=$true)]
[string] $ApiUrl,
[parameter(Position=4, Mandatory=$true)]
[string] $InstanceId
)

try{
  $CurrentDate = Get-Date
  $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh:mm:ss')
  $CurrentDate1 = $CurrentDate.ToString('MM-dd-yyyy')

    $CurrentDate + " User Logged in Successfully in " + $TargetOrganizationURL | Add-Content '.\HygeineLogging.log'
    
 #Region Url

 #Getting BackUp Details
    $backup=Get-CrmInstanceBackups -ApiUrl $Api -Credential $Credential -InstanceId $instaceid    

    $result=$backup | Select-Object -Property Date,Id,Status,CreatedBy,CreatedOn,ExpiresOn,Version,Notes,Label
   
   $CurrentDate + " BackUp retreived Successfully of " +$TargetOrganizationURL | Add-Content '.\HygeineLogging.log'

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

     $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
    $CurrentDate + $_.System.Net.WebException | Set-Content '.\Error.log'
 #    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{        
   $CurrentDate + "Could not find path" | Add-Content '.\HygeineLogging.log'
   $CurrentDate + $_.System.IO.FileNotFoundException | Set-Content '.\Error.log'
   # [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
  
     $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
     $CurrentDate + $_.System.IO.IOException | Set-Content '.\Error.log'
   #   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
  }
    catch
   {
 #  [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Backup Details','Ok','Error')
 
      if(Test-Path '.\HygeineLogging.log'){
       $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
       $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
    }
    else{
      $CurrentDate +  "An Error has occured" | Out-File '.\HygeineLogging.log'
      $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
        
            }
    exit
   }

}
Get-BackUpDetails






# This Function is used get the details of licenses
function Get-LicenceDetails{
param(
[parameter(Position=1, Mandatory=$true)]
[PSCredential]$Credential
)
try{
  $CurrentDate = Get-Date
  $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh:mm:ss')
  $CurrentDate1 = $CurrentDate.ToString('MM-dd-yyyy')

 #Connecting O365
    Connect-MsolService -Credential $Credential


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
   
     $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
     $CurrentDate + $_.System.Net.WebException | Set-Content '.\Error.log'
   #  [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{        
    $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
    $CurrentDate + $_.System.IO.FileNotFoundException | Set-Content '.\Error.log'
    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
      $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
   $CurrentDate + $_.System.IO.IOException | Set-Content '.\Error.log'
      [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
  }
    catch
   {
     $CurrentDate | Add-Content '.\HygeineLogging.log'
    # [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','Licence Details','Ok','Error')
      if(Test-Path '.\HygeineLogging.log'){
        $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
        $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
        
    }
    else{
        $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
        $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
            }
    exit
   }
}

Get-LicenceDetails







#This Function is getting the details of the failed and crashed Plugin
function Get-PluginDetails{
param(
[parameter(Position=1, Mandatory=$true)]
[PSCredential]$Credential, 
[parameter(Position=2, Mandatory=$true)]
[string] $TargetOrganizationURL
)
try{
 #Connecting CRM
   $trgtCRMOrg = Connect-CrmOnline -Credential $Credential -ServerUrl $TargetOrganizationURL

 #Fetch XML on Entity Plugin Details
$fetchXml = @"

<fetch>
  <entity name="plugintypestatistic" >
    <attribute name="plugintypeid" />
    <attribute name="failurecount" />
    <attribute name="failurepercent" />
    <attribute name="executecount" />
    <attribute name="crashcount" />
    <attribute name="crashpercent" />
    <attribute name="averageexecutetimeinmilliseconds" />
    <link-entity name="plugintype" from="plugintypeid" to="plugintypeid" >
      <attribute name="name" alias="PluginName" />
    </link-entity>
  </entity>
</fetch>

"@

$CurrentDate + " Plugin Details retreived Successfully of " +$trgtOrgUrl | Add-Content '.\HygeineLogging.log'

 #Creating Tablubar format to store in CSV File
    $tableEntity = New-Object system.Data.DataTable "PluginStatistics"
    $tblcol1 = New-Object system.Data.DataColumn Date, ([string])
    $tblcol2 = New-Object system.Data.DataColumn PLuginTypeID, ([string])
    $tblcol3 = New-Object system.Data.DataColumn failurecount, ([string])
    $tblcol4 = New-Object system.Data.DataColumn failurepercent, ([string])
    $tblcol5 = New-Object system.Data.DataColumn executecount, ([string])
    $tblcol6 = New-Object system.Data.DataColumn crashcount, ([string])
    $tblcol7 = New-Object system.Data.DataColumn crashpercent, ([string])
    $tblcol8 = New-Object system.Data.DataColumn averageexecutetimeinmilliseconds, ([string])
    $tblcol9 = New-Object system.Data.DataColumn PluginName, ([string])
    $tableEntity.columns.add($tblcol1)
    $tableEntity.columns.add($tblcol2)
    $tableEntity.columns.add($tblcol3)
    $tableEntity.columns.add($tblcol4)
    $tableEntity.columns.add($tblcol5)
    $tableEntity.columns.add($tblcol6)
    $tableEntity.columns.add($tblcol7)
    $tableEntity.columns.add($tblcol8)
    $tableEntity.columns.add($tblcol9)

    $FetchResult = Get-CrmRecordsByFetch -conn $trgtCRMOrg -Fetch $fetchXml -AllRows
   # $solutionEntity = New-Object System.Collections.Generic.List[Guid]

    $FetchResult.CrmRecords | ForEach-Object {
     $tblrow = $tableEntity.NewRow()
          $tblrow.Date = $CurrentDate
          $tblrow.PLuginTypeID = $_.PLuginTypeID
          $tblrow.failurecount = $_.failurecount
          $tblrow.failurepercent = $_.failurepercent
          $tblrow.executecount = $_.executecount
          $tblrow.crashcount = $_.crashcount
          $tblrow.crashpercent = $_.crashpercent
          $tblrow.averageexecutetimeinmilliseconds = $_.averageexecutetimeinmilliseconds
          $tblrow.PluginName = $_.PluginName
          
          $tableEntity.Rows.Add($tblrow)        
    }

    
    #Exporting Data in CSV File 
    $tableEntity | Export-Csv -Path '.\PluginDetails.csv' -Append

    $CurrentDate + " The Plugin Data has been exported." | Add-Content '.\HygeineLogging.log'
    "           " | Add-Content '.\PluginDetails.csv'

  #  [System.Windows.MessageBox]::Show('Operation Successfull!! Please See Target File for information',' Details','Ok','Information')

    $CurrentDate + " Plugin Module run Successfully Successfully" | Add-Content '.\HygeineLogging.log'

  }
  catch [System.Net.WebException]
  {
   
     $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
     $CurrentDate + $_.System.Net.WebException | Set-Content '.\Error.log'
    # [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information',' Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{        
    $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
    $CurrentDate + $_.System.IO.FileNotFoundException | Set-Content '.\Error.log'
   # [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information',' Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
       $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
       $CurrentDate + $_.System.IO.IOException | Set-Content '.\Error.log'
   #   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information',' Details','Ok','Error')
  }
    catch
   {
 #  [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information',' Details','Ok','Error')
  
      $CurrentDate | Add-Content '.\HygeineLogging.log'
      if(Test-Path '.\HygeineLogging.log'){
        $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
        $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
        
    }
    else{
        $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
        $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
            }
    exit
   }

  }

Get-PluginDetails









#This Function is getting the details of the waiting and failed System Job
function Get-Systemjob{
param(
[parameter(Position=1, Mandatory=$true)]
[PSCredential]$Credential, 
[parameter(Position=2, Mandatory=$true)]
[string] $TargetOrganizationURL
)

try{
  $CurrentDate = Get-Date
  $CurrentDate = $CurrentDate.ToString('MM-dd-yyyy hh:mm:ss')
  $CurrentDate1 = $CurrentDate.ToString('MM-dd-yyyy')
 #User Logging Credentials
    $trgtOrgUrl = "https://technoyash.crm8.dynamics.com"

 #Connecting CRM
   $trgtCRMOrg = Connect-CrmOnline -Credential $Credential -ServerUrl $TargetOrganizationURL

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
   Send-Email $locBack $locLicense $locSystem


  }

 #catching exceptions
  catch [System.Net.WebException]
  {
       $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
       $CurrentDate + $_.System.Net.WebException | Set-Content '.\Error.log'
 #    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
     
  }
  catch [System.IO.FileNotFoundException],[System.IO.DirectoryNotFoundException]
{      
     $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
     $CurrentDate + $_.System.IO.FileNotFoundException | Set-Content '.\Error.log'
    [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
}
  catch [System.IO.IOException]
  {
     $CurrentDate + "An Error has occured" | Add-Content '.\HygeineLogging.log'
     $CurrentDate + $_.System.IO.IOException | Set-Content '.\Error.log'
   #   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
  }
    catch
   {
   [System.Windows.MessageBox]::Show('An error occured. Please See Log File For More Information','getsystemjob Details','Ok','Error')
  
      if(Test-Path '.\HygeineLogging.log'){
       $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
       $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'  
    }
    else{
        $CurrentDate + "An Error has occured" | Out-File '.\HygeineLogging.log'
      $CurrentDate + $_.Exception.Message | Set-Content '.\Error.log'
        
            }
    exit
   }
  }

 Get-Systemjob
