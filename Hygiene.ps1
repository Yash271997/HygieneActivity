function GetList() {
    $account_qry = New-Object Microsoft.Xrm.Sdk.Query.QueryExpression("account")
    $account_qry.ColumnSet = New-Object Microsoft.Xrm.Sdk.Query.ColumnSet("name", "accountnumber","statecode", "statuscode")
    #$account_qry.Criteria.AddCondition("industrycode", [Microsoft.Xrm.Sdk.Query.ConditionOperator]::NotNull);
    #$account_qry.Criteria.AddCondition("accountnumber", [Microsoft.Xrm.Sdk.Query.ConditionOperator]::Equal, 'LANG016');
    $account_qry.AddOrder("name", [Microsoft.Xrm.Sdk.Query.OrderType]::Ascending);

    ## RetrieveMultiple returns a maximum of 5000 records by default.
    ## If you need more, use the response's PagingCookie.
    $account_response = $crm_service.RetrieveMultiple($account_qry)

    ## create a dictionary that maps each account key is (name, accountnumber) & id is accountid:
    $account_dic = @{ }
    $account_response.Entities | ForEach-Object {
        $name = $_.Attributes["name"];
        $accountnumber = $_.Attributes["accountnumber"];
        $industrycode = $_.Attributes["industrycode"].value
        $statecode = $_.Attributes["statecode"].value
        $statuscode = $_.Attributes["statuscode"].value
        $parentaccountid = $_.Attributes["parentaccountid"] ## $parentaccountid.Id
        $ownerid = $_.Attributes["ownerid"]
        $value = ("{0}|{1}|{2}|{3}" -f @($name, $accountnumber, $statecode, $statuscode))
        $account_dic.Add($_.Id, $value)
    }

    ## account_dic contact values
    $account_dic.GetEnumerator() | ForEach-Object{
        $account = 'Account {0} -  {1}' -f $_.key, $_.value
        Write-Output $account
    }

    <#
        Account efe7e3ac-9fc5-e711-80f7-0050568c5bd0 -  SQL Server|LANG200|18|C#|UAT|0|1
        Account dc57a4c1-8bc5-e711-80f7-0050568c5bd0 -  PowerShell|LANG016|8||DNS|1|2
    #>
}



PS C:\> $result = Get-CrmRecords -EntityLogicalName asyncoperation -FilterAttribute statuscode -FilterOperator eq "10" -Fields name,statuscode


foreach($i in $result.CrmRecords){ Echo $i.name;$i.statuscode}

$result | Export-Csv -Path C:\Yash\Data.csv

Add-Content -Path C:\Yash\sys.csv -Value '"Name","Sattus"'

$p | foreach{Add-Content -Path C:\Yash\account2.csv -Value $_}


PS C:\> $p | Export-Csv -Path C:\Yash\account.csv

PS C:\> $p | Export-Csv -Path C:\Yash\account1.csv -NoTypeInformation

PS C:\> Import-Csv -Path C:\Yash\account.csv
$result = Get-CrmRecords -EntityLogicalName account -FilterAttribute name -FilterOperator like -FilterValue 'A%' -Fields name


 $result = Get-CrmRecords -EntityLogicalName asyncoperation -FilterAttribute statuscode -FilterOperator eq "10" -Fields name,statuscode

 $p | foreach{Add-Content -Path C:\Yash\account3.csv -Value $_}

 
Add-Content -Path C:\Yash\account3.csv -Value '"Name","Trust ID"'

$result = Get-CrmRecords -EntityLogicalName asyncoperation -Fields name,statuscode

Add-Content -Path C:\Yash\job_14_june.csv -Value '"Name","Status"'

PS C:\> $p=$result

PS C:\> $result = Get-CrmRecords -EntityLogicalName asyncoperation -FilterAttribute statuscode -FilterOperator eq "30" -Fields name,statuscode,friendlymessage

PS C:\> $p = foreach($i in $result.CrmRecords){ Echo $i.name;$i.statuscode}

PS C:\> $p | foreach{Add-Content -Path C:\Yash\jon_14_june.csv -Value $_}

