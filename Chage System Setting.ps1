#User Logging Credentials
  $trgtOrgUrl = "https://skyhightechno.crm8.dynamics.com"
  $trgtUserName = "yash@skyhightechno.onmicrosoft.com"
  $trgtPass = ConvertTo-SecureString "qwerty@1234" -AsPlainText -Force
  $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $trgtUserName, $trgtPass
 
$conn = Connect-CrmOnline -Credential $cred -ServerUrl $trgtOrgUrl


$r = Get-CrmRecords -conn $conn -EntityLogicalName organization -Fields organizationid
$a=foreach($i in $r.CrmRecords){ $i.organizationid.Guid}

$FilePath = "V:\POwer Shell Practice\org.xlsx"
$objExcel = New-Object -ComObject Excel.Application
$WorkBook = $objExcel.Workbooks.Open($FilePath)
$Sheet = $WorkBook.sheets.item("ChangeSetting")

$row = 3
$column = 1

for($i=0; $i -le 595; $i++) {
  if($Sheet.Cells.Item($row,$column).Text -eq "Change"){      
     $SettingName = $Sheet.Cells.Item($row-2,$column).Text
     $SettingName
     $SettingValue = $Sheet.Cells.Item($row-1,$column).Text
     $SettingValue
     If($SettingValue.Equals('Yes')){
     Write-Host YES
     Set-CrmRecord -conn $conn -EntityLogicalName organization $a @{$SettingName=$true}
     $column = $column + 1
     continue
        }
     If($SettingValue.Equals('No')){
     Write-Host NO
     Set-CrmRecord -conn $conn -EntityLogicalName organization $a @{$SettingName=$false}
     $column = $column + 1
     continue
   }
     if($SettingValue -match "^[0-9]*$"){
        $cv=[int]$SettingValue
        Set-CrmRecord -conn $conn -EntityLogicalName organization $a @{$SettingName=$cv}
        $column = $column + 1
        continue
     }
     else{
        $cv=[string]$SettingValue
        Set-CrmRecord -conn $conn -EntityLogicalName organization $a @{$SettingName=$cv}
        $column = $column + 1
        continue
     }
 }
    $column = $column + 1
 }
 (Get-CrmRecords -conn $conn -Entitylogicalname organization -Fields *).CrmRecords | Export-Excel 'V:\POwer Shell Practice\org.xlsx' -Append

 Start-Sleep 2
 $objExcel.Quit()

 #Setting can chnage via PowerShell
 #Set-CrmSystemSettings -conn $conn -AllowUsersSeeAppdownloadMessage lk -CurrencyDisplayOption js -QuickFindRecordLimitEnabled ujh -AMDesignator -IsAuditEnabled uh -IsUserAccessAuditEnabled sf -RequireApprovalForUserEmail sf -RequireApprovalForQueueEmail f -MaxUploadFileSize jj -EnableSmartMatching ujh

