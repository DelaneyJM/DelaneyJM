try {
    $pri = @{
        Client1 = "SRV-CLIENT1-CM1"
        Client2 = "SRV-CLIENT2-CM1"
        Client3 = "SRV-CLIENT3-CM1"
    }
    if($env:USERDNSDOMAIN -eq "Client4.ca") {
        if($env:ComputerName -eq "SRV-CLIENT4-CM1") {
            $pri['Client4'] = "DTP-SCCM-V000"
        } else {
            $pri['Client4'] = "SRV-CLIENT4-CM2"
        }
    }
    $cs = Get-CimSession -Name SCCM -ErrorAction Stop
} catch {
    if(($pri.GetEnumerator() | Select-Object -ExpandProperty Value) -notcontains $env:ComputerName) {
        ($pri.GetEnumerator()).Where({
            $env:USERDNSDOMAIN -match $_.Name
        }) | Select-Object -ExpandProperty Name -OutVariable ToSession > $null
        $SessionParams = @{
            Name = "SCCM"
            ComputerName = $ToSession
            SessionOption = New-CimSessionOption -Protocol Dcom
        }
        $cs = New-CimSession @SessionParams
    }
} finally {
    # We know what environment we are in.  Set the base namespace
    $nsi = $pri.GetEnumerator() |
        Where-Object { $env:USERDNSDOMAIN -match $_.name } |
            Select-Object -ExpandProperty Name
    switch ($nsi) {
        "Client1" { $ns = "root\sms\site_abc"; }
        "Client2"  { $ns = "root\sms\site_def"; }
        "Client3" { $ns = "root\sms\site_ghi"; }
        "Client4"   {
            if($env:ComputerName -eq "DTP-SCCM-V000") {
                $ns = "root\sms\site_jkl";
            } else {
                $ns = "root\sms\site_mno"
            }
        }
        default { $ns = $null;               }
    }

    # We need know if we are running this on a primary because
    # otherwise, Get-CimInstance will fail to connect properly.
    if(-not $cs) {
        $OnPrimary = $true
    } else {
        $OnPrimary = $false
    }
}

# Build the OS requirement object
$fs = "localizeddisplayname LIKE '%Windows%11%' AND modelname NOT LIKE '%ARM%'"
$CimParams = @{
    Namespace = $ns
    ClassName = "SMS_ConfigurationPlatform"
    Filter = $fs
}
if(-not $OnPrimary) {
    $CimParams['CimSession'] = $cs
}
$Win11 = Get-CimInstance @CimParams

$pspcappq = "
SELECT DISTINCT
    app.localizeddisplayname
    ,app.objectpath
    ,app.modelname
FROM sms_applicationlatest AS app
JOIN sms_applicationassignment AS aa
    ON app.modelid = aa.appmodelid
JOIN sms_collection AS col
    ON aa.targetcollectionid = col.collectionid
WHERE
    col.objectpath LIKE '/Client1%Collections%' AND
    app.isdeployed = 1 AND
    app.isexpired = 0
"

# Get Application list
$CimParams = @{
    Namespace = $ns
    Query = $pspcappq -replace "\s\s+", " "
}
if(-not $OnPrimary) {
    $CimParams['CimSession'] = $cs
}
$AppList = Get-CimInstance @CimParams

$ToBeProcessed = foreach($App in $AppList) {
    $PSDefaultParameterValues['disabled'] = $true
    $tmpapp = $app | Get-CimInstance
    $PSDefaultParameterValues.Remove("disabled")
    $osxml = [xml]($tmpapp.SDMPackageXML)
    $osxml.AppMgmtDigest.DeploymentType |
        Select-Object @{l="ApplicationName";e={$app.localizeddisplayname}},
            @{l="ModelName";e={$app.modelname}},
            @{l="ApplicationPath";e={$app.objectpath}},
            @{l="ApplicationXML";e={$tmpapp.sdmpackagexml}},
            @{l="DTName";e={$_.title."#text"}},
            @{l="HasOSRequirement";e={$null -ne $_.requirements.rule.operatingsystemexpression}},
            @{l="HasWin10OSRequirement";e={
                $rules = $_.requirements.rule.operatingsystemexpression.operands.ruleexpression.ruleid
                if([regex]::matches($rules,"Win.*10.*").groups) { 
                    $true 
                } else { 
                    $false 
                } 
            }},
            @{l="HasWin11OSRequirement";e={
                $rules = $_.requirements.rule.operatingsystemexpression.operands.ruleexpression.ruleid
                if([regex]::matches($rules,"Win.*11.*").groups) { 
                    $true 
                } else { 
                    $false 
                } 
            }}
}
$ToBeFinalized = $ToBeProcessed.Where({
    $_.hasosrequirement -and 
    $_.haswin10osrequirement -and 
    -not $_.haswin11osrequirement
}) | Sort-Object ApplicationPath, ApplicationName

$objpath = ""
foreach($app in $ToBeFinalized[0..24]) {
    if($objPath -eq "") {
        Write-Verbose "Starting in folder $($app.applicationpath)..." -Verbose
    } else {
        if($objPath -ne $app.applicationpath) {
            Write-Verbose "Exiting folder $($objPath)..." -Verbose
            Write-Verbose "Starting in folder $($app.applicationpath)..." -Verbose
        }
    }
    $objpath = $app.applicationpath
    $vbm = "$("Modifying DT $($app.dtname) on $($app.applicationname)")"
    Write-Verbose -Message $vbm -Verbose
    
    $job = Start-Job -Name "AddWin11Req - $($app.ApplicationName)|$($app.dtname)" -ArgumentList $app, $Win11 -ScriptBlock {
        param(
            [object]$FinalApp,
            [object]$win11type
        ) 
        $JoinParams = @{
            Path = Split-Path -path $env:SMS_ADMIN_UI_PATH -parent
            ChildPath = "configurationmanager.psd1" 
            Resolve = $true
        }
        $sccmmodpath = Join-Path @JoinParams
        Import-Module $sccmmodpath
        $xml = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::DeserializeFromString(
            $FinalApp.applicationxml,
            $true
        )
        $DeploymentType = $xml.DeploymentTypes | 
            Where-Object { $_.title -eq $finalapp.dtname }
        $requirement = $DeploymentType.Requirements | 
            Where-Object { ($_.Expression.GetType()).Name -eq "OperatingSystemExpression"}
        if($requirement.Name.Normalize() -notmatch "All Windows 11") { 
            foreach($operand in $Win11type) {
                $requirement.Expression.Operands.Add("$($operand.modelname)")
                $requirement.Name = [regex]::replace(
                    $requirement.Name,
                    '(?<=Operating system One of {)(.*)(?=})', "`$1, $($operand.localizeddisplayname)"
                )
            }
            $null = $DeploymentType.Requirements.Remove($requirement)
            $requirement.RuleId = "Rule_$([guid]::NewGuid())"
            $null = $DeploymentType.Requirements.Add($requirement)
        }
        $UpdatedXML = [Microsoft.ConfigurationManagement.ApplicationManagement.Serialization.SccmSerializer]::SerializeToString(
            $XML,
            $True
        )
        
        #Push-Location CM1:/
        #$TempAppObj = Get-CMApplication -Modelname $FinalApp.ModelName
        #$TempAppObj.sdmpackagexml = $UpdatedXML
        #$TempAppObj.Put()
        #$TempAppObj | Set-CMApplication
        #Pop-Location  
    } 
    Write-Verbose "Background job $($job.Name) has started to add Win11 OS requirements" -Verbose
    $job | Wait-Job > $null
}
