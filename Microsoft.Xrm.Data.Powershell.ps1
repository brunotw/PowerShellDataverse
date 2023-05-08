#Connect to CRM Online through Client

if ($null -eq $service -or $false -eq $service.IsReady ) {
    Write-Host "Service was null. Generating new Connection" -ForegroundColor Green
    $service = Connect-CrmOnline -ServerUrl "https://devbw.crm4.dynamics.com"`
        -ClientSecret "Emh8Q~jyJYqq3I~kyavapztV0AHuREtF_-CEpdgL"`
        -OAuthClientId "409c6885-694f-4b09-a159-90fdfc4d57b2"`
        -OAuthRedirectUri "https://tempuri.org"

        #ByPass Custom Business Logic
        $service.BypassPluginExecution = $true;
}
else {
    Write-Host "Using existing connection." -ForegroundColor Green
}

if ($service.IsReady) {
    Write-Host "Successfully connected to:" $service.CrmConnectOrgUriActual -ForegroundColor Green
}
else {
    Write-Host "Couldn't connect to CRM." -ForegroundColor Red
    Write-Host "Last Crm Error:" $service.LastCrmError -ForegroundColor Red
    Write-Host "Last Crm Exception:" $service.LastCrmException -ForegroundColor Red
}

#Execute WhoAmI Request
write-host "Executing WhoAmI Request" -ForegroundColor Green
$request = New-Object Microsoft.Crm.Sdk.Messages.WhoAmIRequest
$response =  $service.Execute($request);

write-host "Connected with user:" $response.UserId -ForegroundColor Green
