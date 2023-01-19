Function Get-GOPDKTPLenovoFriendlyModelFamilyName {
    <#
        .SYNOPSIS
            Allows admins to search for the friendly model name of a Lenovo
            workstation.
        .DESCRIPTION
            Script uses Invoke-RestMethod cmdlet to retrieve a list of 
            workstations directly from the vendor that is actively maintained.
            Script then compares a 4 character model code supplied by the user 
            and returns the friendly model name family for that particular code.
        .EXAMPLE
            Get-GOPDKTPLenovoFriendlyModelFamilyName -ModelCode 20ME

            ModelCode FriendlyName
            --------- ------------
            20ME      ThinkPad P1

        .EXAMPLE
            "20ME","20L8","20T1" | Get-GOPDKTPLenovoFriendlyModelFamilyName
            
            ModelCode FriendlyName
            --------- ------------
            20ME      ThinkPad P1
            20L8      ThinkPad T480s
            20T1      ThinkPad T14s
        .INPUTS
            ModelCode
                ModelCode is a 4 character parameter that cannot be null or 
                empty, and is a required parameter for this function.
        .OUTPUTS
            Returns a PSObject that lists the user supplied modelcode and the
            corresponding model name family for that code.
        .NOTES
            -Updated by Mike Delaney, 18/01/2023
              -Added pipeline support
              -Added standard parameters via CmdletBinding
              -Changed output to be a PSObject
              -Ensured the request via RestMethod can make an SSL/TLS connection to
                the file hosted at Lenovo.
            -Original version by Damien Van Robaeys (@syst_and_deploy)
              https://www.systanddeploy.com/2023/01/get-list-uptodate-of-all-lenovo-models.html?m=0
        .ROLE
            None
#>
    [CmdletBinding(SupportsShouldProcess)]
    Param(
        [Parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true,
            ValueFromRemainingArguments=$false,
            Position=0
        )]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
        [ValidateLength(4,4)]
        [string]$ModelCode
    ) 
    BEGIN {
        $url = "download.lenovo.com/bsco/schemas/list.conf.txt"
        try {
          $list = (Invoke-RestMethod -Uri $url).Split("`n")
        } catch {
          try {
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            $list = (Invoke-RestMethod -Uri $url).Split("`n")
          } catch {
            throw $_
          }
        } 
    }
    PROCESS {
        if($PSCmdlet.ShouldProcess($url,"GetModelList")) {
            $ModelCode | Select-Object @{l="ModelCode";e={ $_ }},
                @{l="FriendlyName";e={
                    $list.where({ ($_.Contains($ModelCode)) }).split("(")[0]
                }}
        }
    }
}
