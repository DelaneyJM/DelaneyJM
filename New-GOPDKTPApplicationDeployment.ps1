Function Get-GOPDKTPValidDeployTargets {
    [CmdletBinding()]
    PARAM()
    @("Users","Devices","Both")
}
Function Get-GOPDKTPApplicationPreviousVersion {
    [CmdletBinding()]
    PARAM(
        [string]$ApplicationName
    )
    $SiteCodeParams = @{
        CimSession = Get-CimSession -Name SCCM
        Namespace = "root\sms"
        ClassName = "sms_providerlocation"
        Filter = "providerforlocalsite = 1"
    }
    $SiteCode = (Get-CimInstance @SiteCodeParams).SiteCode

    $NewAppParams = @{
        ClassName = "sms_applicationlatest"
        CimSession = Get-CimSession -Name SCCM
        Namespace = "root\sms\site_" + $sitecode
        Filter = "localizeddisplayname = '$($AppName)'"
    }
    $app = Get-CimInstance @NewAppParams

    $PreviousVersions = $NewAppParams
    $fs = "objectpath = '$($app.objectpath)' AND "
    $fs += "localizeddisplayname != '$($AppName)' AND "
    $fs += "SUBSTRING(localizeddisplayname,1,6) = '" + $AppName.Remove(6) + "'"
    $PreviousVersions['Filter'] = $fs
    (Get-CimInstance @PreviousVersions).localizeddisplayname
}
Function New-GOPDKTPApplicationDeployment {
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("SSC", "PSPC", "CSPS", "INFC")]
        [string]$Department,

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("RATE", "Production")]
        [string]$Environment,

        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ArgumentCompleter({
            param(
                $Command, $Parameter, $WordToComplete,
                $CommandAst, $FakeBoundParams
            )
            Get-GOPDKTPValidDeployTargets
        })]
        [ValidateScript({
            $_ -in $(Get-GOPDKTPValidDeployTargets)
        })]
        [string]$TargetDeploy = "Users",

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            $AppName = $_
            $SiteCodeParams = @{
                CimSession = Get-CimSession -Name SCCM
                Namespace = "root\sms"
                ClassName = "sms_providerlocation"
                Filter = "providerforlocalsite = 1"
            }
            $SiteCode = (Get-CimInstance @SiteCodeParams).SiteCode

            $ValidAppParams = @{
                ClassName = "sms_applicationlatest"
                CimSession = Get-CimSession -Name SCCM
                Namespace = "root\sms\site_" + $sitecode
                Filter = "localizeddisplayname = '$($AppName)'"
            }
            $ValidApp = (Get-CimInstance @ValidAppParams).localizeddisplaynname
            $errstr = "Release Distribution number not present for $($AppName)."
            $errstr += "  Update application name and try again"
            if($ValidApp) {
                $r = "R[0-9]{1,3}D[0-9]{1,3}"
                $Verify = $ValidApp.localizeddisplayname -cmatch $r
                if($Verify) {
                    $true
                } else {
                    New-Object System.Management.Automation.ErrorRecord(
                        (New-Object System.MissingFieldException($errstr)),
                        'InvalidApplicationName',
                        [Management.Automation.ErrorCategory]::InvalidData,
                        $ApplicationName
                    ) -OutVariable errorRecord
                    throw $errorRecord
                }
            } else {
                $errstr = $errstr -creplace "Update\ application","Verify"
                New-Object System.Management.Automation.ErrorRecord(
                    (New-Object System.MissingFieldException($errstr)),
                    'InvalidApplicationName',
                    [Management.Automation.ErrorCategory]::InvalidData,
                    $ApplicationName
                ) -OutVariable errorRecord
                throw $errorRecord
            }
        })]
        [string]$ApplicationName,

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Vendor,

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$ProductName,

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$OPI,

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet(
            "Win 7", "Win 7/8.1", "Win 7/8.1/10",
            "Win 7/10", "Win 8.1", "Win 8.1/10",
            "Win 10","Win 11", "Win10/11")]
        [string]$SupportedOS,

        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [string]$Dependencies = "N/A",

        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ArgumentCompleter({
            param(
                $Command, $Parameter, $WordToComplete,
                $CommandAst, $FakeBoundParams
            )
            Get-GOPDKTPApplicationPreviousVersion
        })]
        [ValidateScript({
            $_ -in $(Get-GOPDKTPApplicationPreviousVersion)
        })]
        [string]$PreviousVersion,

        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [switch]$RemovePreviousVersion = $false,

        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("MigrateUsers", "DoNotMigrateUsers", "NotSpecified")]
        [string]$ADPreviousVersionAction = "NotSpecified",

        [Parameter(
            Mandatory=$false,
            ValueFromPipeline=$false,
            ValueFromPipelineByPropertyName=$false,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [switch]$RequireReboot = $false
    )

    BEGIN {
        # Create a StringBuilder object to build the text that appears in the
        # Notes field in Active Directory.  Append the necessary lines and
        # redirect the output to null since we don't need to see it.
        $info = New-Object -TypeName System.Text.StringBuilder
        $vbm = New-Object -TypeName System.Text.StringBuilder

        # Determine who is creating this deployment.
        $engineer = switch -regex ($env:username.tolower()) {
            "\bsavary.*\b"   { "Michel Savary"; break }
            "\bdelaney.*\b"  { "Mike Delaney";  break }
            "\bcampbell.*\b" { "Ian Campbell";  break }
            "\bhunton.*\b"   { "Kevin Hunton";  break }
            "\bsilva.*\b"    { "Robert Silva";  break }
            "\bchapman.*\b"  { "Scott Chapman"; break }
            default          { $env:username;   break }
        }
        if($engineer) {
            $ProcessedDeployments = Get-Job
            $merid = ""
            if(-not $ProcessedDeployments) {
                # No Deployments created today
                if((Get-Date).Hour -gt 11) {
                    $merid = "Afternoon"
                } else {
                    $merid = "Morning"
                }
            } else {
                # At least 1 deployment record created
                $TodaysDeploys = $ProcessedDeployments |
                    Where-Object {
                        ((Get-Date) - ($_.psbegintime)).totaldays -lt 1
                    }
                if(-not $TodaysDeploys) {
                    # Nothing deployed yet today
                    if((Get-Date).Hour -gt 11) {
                        $merid = "Afternoon"
                    } else {
                        $merid = "Morning"
                    }
                }
            }
            if($merid) {
                $vbm.Append("[GOPDKTP] { BEGIN } Good $($merid), ") > $null
                $vbm.Append("$($engineer)") > $null
            } else {
                $vbm.Append("[GOPDKTP] { BEGIN } Thank you for ") > $null
                $vbm.Append("giving me another job to do $($engineer)!") > $null
            }
        }
        Write-Verbose -Message $vbm.ToString() -Verbose
        $vbm.Clear() > $null

        # Flag the AD group for workstations only if it's going to be used for
        # device collection query that enumerates the AD group membership.
        if($PSBoundParameters.ContainsKey("TargetDeploy")) {
            if($TargetDeploy -eq "Devices") {
                $vbm.Append("[GOPDKTP] { BEGIN } Adding workstations ") > $null
                $vbm.Append("only are to be added to new AD group") > $null
                Write-Verbose -Message $vbm.ToString() -Verbose
                $l = "WORKSTATIONS ONLY. Populates a device collection in SCCM."
                $l += " Users will be ignored."
                $info.AppendLine("$($l)") > $null
                $info.AppendLine(" ") > $null
                $vbm.Clear() > $null
            }
        }

        # Add the INFOWEB product name.
        $vbm.Append("[GOPDKTP] Add product name to Notes field.") > $null
        Write-Verbose -Message $vbm.ToString() -Verbose
        $info.AppendLine("INFOMAN: $($ProductName)") > $null
        $vbm.Clear() > $null

        # Identify who the OPI is.
        $vbm.Append("[GOPDKTP] { BEGIN } Add OPI to Notes field.") > $null
        Write-Verbose -Message $vbm.ToString() -Verbose
        $info.AppendLine("Application Owner (OPI): $($OPI)") > $null
        $vbm.Clear() > $null

        # Add which OS versions are supported for this release
        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Add supported OS") > $null
        Write-Verbose -Message $vbm.ToString() -Verbose
        $info.AppendLine("Platform Supported: $($SupportedOS)") > $null
        $vbm.Clear() > $null

        # Application dependencies are an optional parameter.  If it's been
        # defined, add it to the notes field.  Otherwise, use the default
        # parameter value of N/A.
        if($PSBoundParameters.ContainsKey("Dependencies")) {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding application ") > $null
            $vbm.Append("dependencies / prerequisites to Notes field") > $null
        } else {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: No application ") > $null
            $vbm.Append("dependencies / prerequisites defined. Adding") > $null
            $vbm.Append("N/A to Notes field") > $null
        }
        Write-Verbose -Message $vbm.ToString() -Verbose
        $vbm.Clear() > $null
        $info.AppendLine(
            "Application Dependencies/Pre-Requisites: $($Dependencies)"
        ) > $null

        # PreviousVersion is an optional parameter.  If it's been
        # defined, add it to the notes field.  Otherwise, use the default
        # parameter value of N/A.
        if($PSBoundParameters.ContainsKey("PreviousVersions")) {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Add previous versions") |
                Out-Null
            $vbm.Append(" to Notes field") > $null
        } else {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: No previous versions.") |
                Out-Null
            $vbm.Append("Adding N/A to Notes field") > $null
        }
        Write-Verbose -Message $vbm.ToString() -Verbose
        $info.AppendLine(
            "Previous Version(s): $($PreviousVersions)"
        ) > $null

        # RemovePreviousVersion is an optional parameter.  If it's been defined,
        # check what the value is, and whether we need to migrate the previous
        # group's membership.
        if($PSBoundParameters.ContainsKey("RemovePreviousVersion")) {
            if($PSBoundParameters.ContainsKey("ADPreviousVersionAction")) {
                switch($ADPreviousVersionAction) {
                    "MigrateUsers" {
                        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding ") > $null
                        $vbm.Append("that previous version is to be ") > $null
                        $vbm.Append("removed from production and the ") > $null
                        $vbm.Append("users of the AD group should be ") > $null
                        $vbm.Append("migrated to a new group.") > $null
                        break
                    }
                    "DoNotMigrateUsers" {
                        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding ") > $null
                        $vbm.Append("that previous version is to be ") > $null
                        $vbm.Append("removed from production and the ") > $null
                        $vbm.Append("users of the AD group will not be") > $null
                        $vbm.Append("migrated to new group.") > $null
                        break
                    }
                    "NotSpecified" {
                        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding ") > $null
                        $vbm.Append("that previous version is to be ") > $null
                        $vbm.Append("removed from production but no ") > $null
                        $vbm.Append("action for the previous version's") > $null
                        $vbm.Append("AD group members was specified") > $null
                        break
                    }
                }
            }
        } else {
            if($PSBoundParameters.ContainsKey("ADPreviousVersionAction")) {
                switch($ADPreviousVersionAction) {
                    "MigrateUsers" {
                        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding ") > $null
                        $vbm.Append("that previous version is not to ") > $null
                        $vbm.Append("be removed from production and ") > $null
                        $vbm.Append("the users of the AD group should ") > $null
                        $vbm.Append("be migrated to a new group.") > $null
                        break
                    }
                    "DoNotMigrateUsers" {
                        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding ") > $null
                        $vbm.Append("that previous version is not to ") > $null
                        $vbm.Append("be removed from production and ") > $null
                        $vbm.Append("the users of the AD group will ") > $null
                        $vbm.Append("not be migrated to new group.") > $null
                        break
                    }
                    "NotSpecified" {
                        $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding ") > $null
                        $vbm.Append("that previous version is not to ") > $null
                        $vbm.Append("be removed from production but no ") > $null
                        $vbm.Append("action for the previous version's") > $null
                        $vbm.Append("AD group members was specified") > $null
                        break
                    }
                }
            }
        }
        Write-Verbose -Message $vbm.ToString() -Verbose
        $vbm.Clear() > $null

        $rRegex = '([A-Z\W_]|\d+)(?<![a-z])'
        $rStr = ' $&'
        $ADAction = $ADPreviousVersionAction -creplace $rRegex, $rStr
        if($PSBoundParameters.ContainsKey("RemovePreviousVersion")) {
            $RemAction = "Yes"
        } else {
            $RemAction = "No"
        }
        $l = "Remove Previous Version: $($RemAction), $($ADAction)"
        $info.AppendLine($l) > $null

        # RequireReboot is an optional parameter.  If it's been defined, SCCM
        # will reboot the workstation once the application deployment has been
        # completed and the application has been detected.
        if($PSBoundParameters.ContainsKey("RequireReboot")) {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding a reboot") > $null
            $vbm.Append(" is required for this software install once") > $null
            $vbm.Append(" deployment is complete") > $null
            $info.Append("Reboot: Yes") > $null
        } else {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Adding a reboot") > $null
            $vbm.Append(" is not required for this software install") > $null
            $vbm.Append(" once deployment is complete") > $null
            $info.Append("Reboot: No") > $null
        }
        Write-Verbose -Message $vbm.ToString() -Verbose
        $vbm.Clear() > $null

        # Detect whether we are on the PSPC/SSC network.  If we are, check that
        # we are not adding extra departmental markers to the Active Directory
        # group that will be created
        if($env:USERDNSDOMAIN -match "pwgsc") {
            $coredomain = $true
            if($Department -eq "SSC") {
                if($Environment -eq "RATE") {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: SSC RATE ")> $null
                    $vbm.Append("environment detected.") > $null
                    $grpName = "$($Department) RT "
                } else {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: SSC PROD ")> $null
                    $vbm.Append("environment detected.") > $null
                    $grpName = "$($Department) "
                }
                $deptfilt = "(\- SSC$)|(\ SSC$)"
            } else {
                if($Environment -eq "RATE") {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: PSPC RATE ")> $null
                    $vbm.Append("environment detected.") > $null
                    $grpName = "RT "
                } else {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: PSPC PROD ")> $null
                    $vbm.Append("environment detected.") > $null
                }
                $deptfilt = "(\- PSPC$)|(\ PSPC$)"
            }
            $grpName += ($ApplicationName -Replace $deptfilt,'').trim()
            $grpName = $grpName.trim()
        } else {
            if($env:USERDNSDOMAIN -match "csps") {
                if($Environment -eq "RATE") {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: CSPS RATE ")> $null
                    $vbm.Append("environment detected.") > $null
                    $grpName = "CSPS Rate Testing"
                } else {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: CSPS PROD ")> $null
                    $vbm.Append("environment detected.") > $null
                    $grpName = $ApplicationName
                }
            }
            if($env:USERDNSDOMAIN -match "infr") {
                if($Environment -eq "RATE") {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: INFC RATE ")> $null
                    $vbm.Append("environment specified.") > $null
                    $grpName = "INFC Rate Testing"
                } else {
                    $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: INFC PROD ")> $null
                    $vbm.Append("environment specified.") > $null
                    $grpName = $ApplicationName
                }
            }
        }
        Write-Verbose -Message $vbm.ToString() -Verbose
        $vbm.Clear() > $null

        # Validate that the new AD group name is less than 63 characters.
        # If it isn't, throw an exception indicating the application name in
        # SCCM must be changed.
        if($grpName.Length -gt 63) {
            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Error with group ")> $null
            $vbm.Append("name detected.  The name is longer than ") > $null
            $vbm.Append("63 characters and cannot be created.")
            Write-Verbose -Message $vbm.ToString() -Verbose
            $vbm.Clear()

            $vbm.Append("[GOPDKTP-SCCM] { BEGIN }: Please update the ") > $null
            $vbm.Append("application object in MECM and try again.") > $null
            Write-Verbose -Message $vbm.ToString() -Verbose
            $vbm.Clear()

            $errstr = "AD Group Name length exceeds 63 characters"
            $errstr += "Update application name in SCCM and try again"
            New-Object System.Management.Automation.ErrorRecord(
                (New-Object System.MissingFieldException($errstr)),
                'InvalidApplicationName',
                [Management.Automation.ErrorCategory]::InvalidOperation,
                $ApplicationName
            ) -OutVariable errorRecord > $null
            throw $errorRecord
        }
    }
    PROCESS {
        #$desc = "Created by $engineer, $(Get-Date -format "dd\/MM\/yyyy")"

        # start with the ActiveDirectory portion
        if($coredomain) {
            # We're running this inside PSPC/SSC domain
            $oupath = "DC=ad,DC=pwgsc-tpsgc,DC=gc,DC=ca"
            $oupath = $oupath.Insert(0,"OU=SSC-SPC Administration,")
            $oupath = $oupath.Insert(0,"OU=Resource Groups,OU=GOPDKTP,")
            $oupath = $oupath.Insert(0,"$($Environment) Application Groups,")
            $oupath = $oupath.Insert(0,"OU=$($vendor),OU=$($Department) ")
        } else {
            if($env:USERDNSDOMAIN -match "csps") {
                $oupath = "DC=csps-efpc,DC=com"
            }
            if($env:USERDNSDOMAIN -match "infr") {
                $oupath = "DC=ad,DC=infrastructure,DC=gc,DC=ca"
                # OU=INFC RATE Application Groups,OU=Resource Groups,OU=Groups,OU=GOPDKTP,OU=SSC-SPC Administration,DC=ad,DC=infrastructure,DC=gc,DC=ca
                # OU=INFC Production Application Groups,OU=Resource Groups,OU=Groups,OU=GOPDKTP,OU=SSC-SPC Administration,DC=ad,DC=infrastructure,DC=gc,DC=ca
            }
        }

        # We've built our base search location.  Check that the OU exists in AD
        $OUParams = @{
            LDAPFilter = "(name=$($vendor))"
            SearchBase = $oupath.Split(',',2)[-1]
            SearchScope = "Subtree"
        }
        $ou = Get-ADOrganizationalUnit @OUParams
        if(-not $ou) {
            # Didn't find the OU. Create it.
            $vbm = "[GOPDKTP-SCCM] { PROCESS }: ActiveDirectory OU does"
            $vbm += " not exist for $($vendor) at "
            $vbm += "$($oupath.Split(",",2)[-1])."
            Write-Verbose -Message $vbm.ToString() -Verbose
            $vbm.Clear() > $null

            $vbm = "[GOPDKTP-SCCM] { PROCESS }: This will be created in"
            $vbm += " order to process the deployment"
            Write-Verbose -Message $vbm.ToString() -Verbose
            $NewOUParams = @{
                Name = $oupath.Split(',',2)[0]
                Path = $oupath.Split(',',2)[-1]
                ProtectedFromAccidentalDeletion = $false
            }
            New-ADOrganizationalUnit @NewOUParams > $null
        } else {
            # Found the OU.
            $vbm = "[GOPDKTP-SCCM] { PROCESS }: ActiveDirectory OU does"
            $vbm += " exist for $($vendor) at "
            $vbm += "$($oupath.Split(",",2)[-1])."
            Write-Verbose -Message $vbm.ToString() -Verbose
            $vbm.Clear() > $null

            $vbm = "[GOPDKTP-SCCM] { PROCESS }: Proceeding with creation "
            $vbm += " of the Active Directory group in ."
            Write-Verbose -Message $vbm.ToString() -Verbose
            $vbm.Clear() > $null
        }
    }
}
<#
        $ADParams = @{
            Description = $desc
            GroupCategory = "Security"
            GroupScope = "Global"
            Name = $grpName
            SamAccountName = $grpName
            Path = $oupath.Split(',',2)[-1]
            PassThru = $true
            ErrorAction = "Stop"
            OutVariable = "adgrp"
        }

        if($Environment -eq "PROD") {
                if($PreviousVersion) {
                    if($ADPreviousVersionAction -eq "MigrateUsers") {
                        $PrevDeploys = @{
                            CimSession = Get-CimSession -Name SCCM
                            Namespace = "root\sms\site_cm1"
                            ClassName = "sms_deploymentsummary"
                            Filter = "applicationname = $($PreviousVersion)"
                        }
                        $cols = Get-CimInstance @PrevDeploys
                        if($cols.Count -gt 0) {
                            $sb = New-Object -Type System.Text.StringBuilder
                            $cols.where({
                                $_.collectionname -notmatch "(NeilB)"
                            }).foreach({
                                $sb.Append("SUBSTRING(name,1,17) != '") > $null
                                $sb.Append("'PWGSC-TPSGC-EM\\RT' AND ") > $null
                                $sb.Append("collectionid = '") > $null
                                $sb.Append("$($_.collectionid)'") > $null

                                $fcmParams = @{
                                    CimSession = Get-CimSession -Name SCCM
                                    Namespace = "root\sms\site_cm1"
                                    ClassName = "sms_fullcollectionmembership"
                                    Filter = $sb.ToString()
                                }
                                $prevAdGrps = Get-CimInstance @fcmParams
                                $members = $prevADGrps | ForEach-Object {
                                    $Ldap = "(&(name=$($_.name.remove(0,16)))"
                                    $Ldap += "(!(
                                        member=*)))"
                                    $PrevADParams = @{
                                        LDAPFilter = $ldap
                                        Properties = "member"
                                    }
                                    $m = Get-ADGroup @PrevADParams
                                    if($m.member.Count -gt 0) {
                                        $m.member
                                    }
                                }
                            })
                        }
                    }
                }
            }
            if($Environment -eq "RATE") {
                $ADParams['OtherAttributes'] = @{ info = $info.ToString() }
            }

                try {

                    $PathExists = Get-ADOrganizationalUnit @OUParams 4>$null
                    Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Path $($p) exists in Active Directory" -Verbose
                    if($PathExists) {
                        Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Creating AD group $($grpName) at $($p)" -Verbose
                        New-ADGroup @ADParams 4>$null > $null
                    }
                } catch {
                    try {
                        Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Issue creating AD group $($grpName) at $($p)" -Verbose
                        if(-not $PathExists) {
                            Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Path $($p) does not exist in Active Directory" -Verbose
                            $ouname, $oupath = $p.Split(",",2)
                            $ouname = $ouname.replace("OU=","")
                            $NewOUParams = @{
                                Path = $oupath
                                Name = $ouname
                                ProtectedFromAccidentalDeletion = $false
                            }
                            Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Creating new OU at ($p)" -Verbose
                            New-ADOrganizationalUnit @NewOUParams 4>$null > $null
                            Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Creating AD group $($grpName) at $($p)" -Verbose
                            New-ADGroup @ADParams 4>$null > $null
                        } else {
                            Get-ADGroup -Identity $grpName -OutVariable adgrp > $null 4>$null
                        }
                    } catch {
                        # Something happened when trying to create the AD group and vendor OU.
                        # Verify your account has the necessary permissions and try again.
                        # Throwing the exception that Active Directory returned when trying to create objects.
                        throw $_
                        }
                    }
                }
                Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Created AD group $($grpName) at $($p) successfully" -Verbose
                Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: Beginning search of $($grpName) AD group in SCCM at $(Get-Date -format "dd\/MM\/yyyy HH:mm:ss")"
                Write-Verbose -Message "[GOPDKTP-SCCM] { PROCESS }: A popup notification will appear once the Group Discovery cycle has completed"


            # Create a background job to watch for the new AD group discovery
            # and once the group is discovered, add it to the RATE deployment
            # collection.  This uses it's own runspace within PowerShell and
            # otherwise frees up the script to complete the rest of the tasks
            # for the application deployment.

            Start-Job -Name "$($adgrp.Name)" -ScriptBlock {
                if($env:userdnsdomain -match "pwgsc") {
                    $SessionParams = @{
                        ComputerName = "VAPW-CM1"
                        Name = "SCCM_CIM_JOB"
                        SessionOption = New-CimSessionOption -Protocol Dcom
                        OutVariable = "cs"
                    }

                    New-CimSession @SessionParams > $null
                    $ns = "root\sms\site_cm1"

                    $CimQueryParams = @{
                        CimSession = $cs
                        Namespace = $ns
                        ClassName = "sms_r_usergroup"
                        Filter = "usergroupname = '$((($using:adgrp).Name))'"
                        OutVariable = "grp"
                    }

                    $sd = Get-Date

                    $sw = New-Object -TypeName System.Diagnostics.Stopwatch
                    $sw.Start()
                    Get-CimInstance @CimQueryParams > $null
                    if(-not $grp) {
                        while((-not $grp) -and ($sw.ElapsedMilliseconds -lt 1200000)) {
                            Get-CimInstance @CimQueryParams > $null
                            if($grp) {
                                break
                            } else {
                                Start-Sleep -Seconds 15
                            }
                        }
                    }
                    $sw.Stop()
                    if($grp) {
                        Add-Type -AssemblyName System.Windows.Forms
                        $global:balloon = New-Object System.Windows.Forms.NotifyIcon
                        $path = (Get-Process -id $pid).Path

                        $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
                        $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
                        $balloon.BalloonTipText = "AD Group $($grp.usergroupname) has been synchronized to SCCM."
                        $balloon.BalloonTipTitle = "Attention $($using:Engineer)"
                        $balloon.Visible = $true
                        $balloon.ShowBalloonTip(20000)

                        Start-Sleep -Milliseconds 7500
                        $CimQueryParams['ClassName'] = "sms_collection"
                        $CimQueryParams['Filter'] = "name = '$(($using:adgrp).Name.Replace('RT ',''))'"
                        $CimQueryParams['OutVariable'] = "col"

                        Get-CimInstance @CimQueryParams > $null
                        if($col) {
                            $DirectRuleParams = @{
                                Namespace = $ns
                                ClassName = "sms_collectionruledirect"
                                ErrorAction = "Stop"
                                ClientOnly = $true
                                Property = @{
                                    ResourceClassName = $grp.CimSystemProperties.ClassName
                                    RuleName = "$($grp.Name). Added by $($using:engineer), $(Get-Date -day 2 -Format 'dd\/MM\/yyyy')"
                                    ResourceID = $grp.ResourceID
                                }
                            }
                            $cmrule = New-CimInstance @DirectRuleParams
                            $InvokeParams = @{
                                MethodName = "AddMembershipRule"
                                Arguments = @{ CollectionRule = [ciminstance]$cmrule }
                                ErrorAction = "Stop"
                            }
                            try {
                                $Update = $col | Invoke-CimMethod @InvokeParams
                                $sw.Elapsed |
                                    Select-Object @{l="GroupName";e={$grp.usergroupname}},
                                        @{l="CollectionName";e={$col.Name}},
                                        @{l="SyncCompletedIn";e={($sd.AddMilliseconds($_.TotalMilliseconds) - $sd)}},
                                        @{l="AddRuleReturnCode";e={$Update.ReturnValue}}
                                $global:balloon = New-Object System.Windows.Forms.NotifyIcon
                                $path = (Get-Process -id $pid).Path
                                if($Update.ReturnValue -eq 0) {
                                    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
                                    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
                                    $balloon.BalloonTipText = "Collection $($col.name) has been updated to SCCM."
                                    $balloon.BalloonTipTitle = "Attention $($using:Engineer)"
                                    $balloon.Visible = $true
                                    $balloon.ShowBalloonTip(20000)

                                    Start-Sleep -Milliseconds 7500
                                }
                            } catch {
                                $PlayWav = New-Object system.windows.media.mediaplayer
                                $PlayWav.Open('C:\users\admin-mike-delaney\desktop\directionunclearpleaserepeat_clean.mp3')
                                $PlayWav.play()
                                Remove-Variable -Name PlayWav
                                $errorRecord = New-Object System.Management.Automation.ErrorRecord(
                                    (New-Object System.ArgumentException("Unable to add $($grp.Name) to collection $($ApplicationName)")),
                                    'InvalidApplicationName',
                                    [System.Management.Automation.ErrorCategory]::InvalidOperation,
                                    $ApplicationName
                                )
                                throw $errorRecord
                            }
                        }
                    }
                }
                Get-CimSession -Name SCCM_CIM_JOB | Remove-CimSession
            } > $null

            # Verify vendor has a folder under the departmental user collections
            $FolderParams = @{
                Name = $vendor
                ItemType = "directory"
                OutVariable = "colpath"
            }
            if($env:userdnsdomain -match "pwgsc") {
                if($Department -eq "PSPC") {
                    $FolderParams['Path'] = "CM1:\UserCollection\PSPC User Collections"

                    $VerifyParams = @{
                        ClassName = "SMS_ObjectContainerNode"
                        Filter = "ParentContainerNodeID = 16777475 AND name = '$($vendor)'"
                    }
                } else {
                    $FolderParams['Path'] = "CM1:\UserCollection\SSC User Collections"
                    $VerifyParams = @{
                        ClassName = "SMS_ObjectContainerNode"
                        Filter = "ParentContainerNodeID = 16777476 AND name = '$($vendor)'"
                    }
                }
            }
            $usercolpath = Get-CimInstance @VerifyParams 4>$null
            if(-not ($usercolpath)){
                New-Item @FolderParams 4>$null > $null
            }
            $colpath = $FolderParams['Path'] + "\" + $vendor

            try {
                $ScheduleParams = @{
                    RecurCount = 1
                    RecurInterval = "Days"
                }
                if((Get-Date).Hour -ge 13 -and (Get-Date).Hour -le 17) {
                    $ScheduleParams['Start'] = (Get-Date 4>$null).AddHours(-12)
                } else {
                    $ScheduleParams['Start'] = (Get-Date -Hour $(Get-Random -Minimum 1 -Maximum 5 4>$null) 4>$null)
                }
                $schedule = New-CMSchedule @ScheduleParams 4>$null

                $UserColParams = @{
                    Comment = "Created by $($engineer), $(Get-Date -Format "dd\/MM\/yyyy" 4>$null)"
                    Name = $adgrp.Name.Replace("RT ","")
                    RefreshSchedule = $schedule
                    LimitingCollectionID = "SMS00004"
                    OutVariable = "RetVal"
                }
                New-CMUserCollection @UserColParams 4>$null > $null
            } catch {
                try {
                    $RetVal = Get-CMCollection -Name $adgrp.Name.Replace("RT ","")
                } catch {
                    throw $_
                }
            }

            $RetVal | Move-CMObject -FolderPath $colpath 4>$null

            $DeployParams = @{
                AllowRepair = $true
                DeployAction = "Install"
                DeployPurpose = "Available"
                TimeBaseOn = "LocalTime"
                UserNotification = "DisplayAll"
                Comment = "Created by Mike Delaney, $(Get-Date -format "dd\/MM\/yyyy" 4>$null)"
                AvailableDateTime = Get-Date -hour 7 -Minute 0 -Second 0 -Millisecond 0
                PersistOnWriteFilterDevice = $false
                CollectionName = $Retval.Name
            }

            $deploys = Get-CMApplication -Name $ApplicationName -OutVariable app 4>$null |
                Get-CMApplicationDeployment -CollectionName $Retval.Name 4>$null

            if($deploys.count -gt 0) {
                Write-Verbose -Message '{GOPDKTP-SCCM} [PROCESS] Application has already been deployed to RATE.'
            } else {
                $app | New-CMApplicationDeployment @DeployParams -ErrorAction Stop > $null 4>$null
            }

            # Package size
            $app = Get-CMApplication -Name $ApplicationName
            $xml = [xml]$app.sdmpackagexml
            $sourcepath = $xml.AppMgmtDigest.DeploymentType.Installer.Contents.Content.Location
            $sourcepath = $sourcepath.Split("\",3,[stringsplitoptions]::RemoveEmptyEntries)[-1]

            $app | Select-Object *displayname, @{l="PackageSize";e={
                $pkgsize = (Get-ChildItem -Path $("X:\" + $sourcepath) -recurse |
                    Measure-Object -Property Length -Sum).Sum
                switch ([math]::truncate([math]::log($pkgsize,1024))) {
                    0 {"$pkgsize Bytes"; break}
                    1 {"{0:n2} KB" -f ($pkgsize / 1KB); break}
                    2 {"{0:n2} MB" -f ($pkgsize / 1MB); break}
                    3 {"{0:n2} GB" -f ($pkgsize / 1GB); break}
                    4 {"{0:n2} TB" -f ($pkgsize / 1TB); break}
                    Default {"{0:n2} PB" -f ($pkgsize / 1pb); break}
                }
            }} | Format-Table -AutoSize
        }
    }
}#>