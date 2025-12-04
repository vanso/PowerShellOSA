<#

PowerShellOSA

Copyright (C) 2025 Vincent Anso

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program; if not, write to the Free Software
Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.

#>

#Requires -Version 7.2

if ( -Not $IsMacOS )
{
    Write-Warning "This module only runs on macOS."
    exit 0
}
else
{
    if (Test-Path -LiteralPath "$PSScriptRoot/PowerShellOSA.zip")
    {
        Write-Verbose "Need to unzip PowerShellOSA.app"
        /usr/bin/unzip -qo $PSScriptRoot/PowerShellOSA.zip -d "$PSScriptRoot/"
        /bin/rm $PSScriptRoot/PowerShellOSA.zip
        /bin/rm -R $PSScriptRoot/__MACOSX
    }

    if (Test-Path -LiteralPath "$PSScriptRoot/PowerShellOSAUI.zip")
    {
        Write-Verbose "Need to unzip PowerShellOSAUI.app"
        /usr/bin/unzip -qo $PSScriptRoot/PowerShellOSAUI.zip -d "$PSScriptRoot/"
        /bin/rm $PSScriptRoot/PowerShellOSAUI.zip
        /bin/rm -R $PSScriptRoot/__MACOSX
    }
}

enum PowerShellOSAOutputFormat
{
    PSCustomObject
    JSON
    PLIST
}

function Invoke-OSA
{
    <#
    
    .SYNOPSIS
    Allowing PowerShell to call AppleScript script or file using the PowerShellOSA application through seamless two-way communication.
    
    #>

    [CmdletBinding(DefaultParameterSetName = "Path")]
    param (
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$Script,
        [Parameter(ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [string]$Path,
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [Parameter(ValueFromPipeline = $true)]
        [Object]$InputObject = $null,
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [Object]$Parameters,
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [ValidateScript({ $(/usr/bin/osalang) -ccontains $_ })]
        [String]$Language="AppleScript",
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [ValidateSet("PSCustomObject","JSON", "PLIST", IgnoreCase = $false)]
        [PowerShellOSAOutputFormat]$OutputFormat="PSCustomObject"
    )
    
    if ($env:POWERSHELLOSA_PATH)
    {
        $PowerShellOSA=$env:POWERSHELLOSA_PATH
    }
    else 
    {
        $PowerShellOSA="$PSScriptRoot/PowerShellOSA.app"
    }

    if (-Not (Test-Path -LiteralPath $PowerShellOSA))
    {
        Write-Warning "PowerShellOSA application not found at path $PowerShellOSA"

        return $null
    }

    $PowerShellOSA="$PowerShellOSA/Contents/MacOS/PowerShellOSA"

    Write-Debug $PowerShellOSA

    if ($PSCmdlet.ParameterSetName -eq "Script")
    {
        $source = "--script=`"$Script`""
    }
    else 
    {
        $source = "--file=`"$Path`""
    }
    
    $Parameters = $($Parameters | ConvertTo-Json -Compress -WarningAction SilentlyContinue)

    if ($PSBoundParameters.ContainsKey('InputObject'))
    {
        if ($input)
        {
            $InputObject = $input
        }
    }

    $InputObject = $($InputObject | ConvertTo-Json -Compress -WarningAction SilentlyContinue) 

    Write-Debug $InputObject

    Write-Debug $source

    ($OutputFormat -eq "PSCustomObject") ? ($RawOutput = "JSON") : ($RawOutput = $OutputFormat) | Out-Null

    # --script=text | --file=filename [--input=JSON-formatted string] [--parameters=JSON-formatted string] [--language=OSALanguageName] [--format=JSON[PLIST]] [--quiet]
    $command = "$PowerShellOSA $source --input='$InputObject' --parameters='$Parameters' --language=$Language --format=$RawOutput --quiet"

    Write-Debug $command

    $result = Invoke-Expression $command

    if ($result)
    {
        switch ($OutputFormat) {
            {($_ -eq "JSON") -or ($_ -eq "PLIST")} {  
                $result
            }
            "PSCustomObject" {
                $result | ConvertFrom-Json
            }
        }
    }
}

function  Invoke-JavaScript 
{
    [CmdletBinding(DefaultParameterSetName = "Path")]
    param (
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]$Script,
        [Parameter(ParameterSetName = 'Path')]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({ Test-Path -LiteralPath $_ })]
        [string]$Path,
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [Parameter(ValueFromPipeline = $true)]
        [Object]$InputObject = $null,
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [Object]$Parameters,
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [ValidateSet("JavaScript")]
        [String]$Language="JavaScript",
        [Parameter(ParameterSetName = 'Script')]
        [Parameter(ParameterSetName = 'Path')]
        [ValidateSet("PSCustomObject","JSON", "PLIST", IgnoreCase = $false)]
        [PowerShellOSAOutputFormat]$OutputFormat="PSCustomObject"
    )

    if ($PSCmdlet.ParameterSetName -eq "Script")
    {
        Invoke-OSA -Script $Script -InputObject $InputObject -Parameters $Parameters -Language $Language -OutputFormat $OutputFormat   
    }
    else 
    {
        Invoke-OSA -Path -$Path $InputObject -Parameters $Parameters -Language $Language -OutputFormat $OutputFormat
    }
}

function ConvertTo-Hashtable
{
    <#
    
    .SYNOPSIS
    Convert a PSCustomObject to a Hashtable.

    #>

    param(
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [PSCustomObject]$Object
    )

    $hashtable = @{} 

    $Object.PSObject.Properties | ForEach-Object { $hashtable[$_.Name] = $_.Value }

    $hashtable
}

function ConvertTo-PSCutomObject
{
    <#
    
    .SYNOPSIS
    Convert a Hashtable to a PSCustomObject.

    #>

    param(
        [hashtable]$Hashtable
    )

    New-Object psobject -Property $Hashtable
}

function ConvertFrom-PascalCase 
{
    <#
    
    .SYNOPSIS
    Convert a Pascal Case word to a string.

    #>

    param (
        [string]$String
    )   

    ($String -creplace '([a-z])([A-Z])', '$1 $2').ToLower()
}

function ConvertTo-PascalCase
{
    <#
    
    .SYNOPSIS
    Convert a string to a Pascal Case word.

    #>

    param (
        [string]$String,
        [string]$Delimiter = " "
    )

    # Remove diacritics
    $normalized = $String.Normalize([Text.NormalizationForm]::FormKD);
    $String = $(-join ($normalized.ToCharArray() | Where-Object { [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark }))

    $String = $String.Replace("_"," ")
    $String = $String.Replace("-"," ")
    $String = $String.Replace("."," ")
    $String = $String.Replace("’s","")
    $String = $String.Replace("#","N")
    $String = $(-join ($String.ToCharArray() | Where-Object { -not [Char]::IsPunctuation($_) }))
    $String = [Text.RegularExpressions.Regex]::Replace($String, "\s+", " ")

    $words = $String -split $Delimiter
    $result = ""

    foreach ($word in $words) 
    {
        if ($word.Length -gt 0)
        {
            $capitalized = $word.Substring(0,1).ToUpper() + $word.Substring(1)
        }

        $result += $capitalized
    }

    return $result
}

function ConvertTo-AppleScriptParameter
{
    <#
    
    .SYNOPSIS
    Convert a PowerShell parameter to an AppleScript parameter.

    #>
    
    param(
        $Parameter
    )

    $value = $null

    Write-Debug $Parameter.GetType()

    Write-Debug "$($Parameter) $($Parameter.Key) $($Parameter.Value) $($Parameter.Value.GetType()) $($Parameter.Value.GetType().BaseType)"

    if ($Parameter.Value.GetType() -eq [switch])
    {
        $value = $($Parameter.Value) ? "with $(ConvertFrom-PascalCase $Parameter.Key)" : "without $(ConvertFrom-PascalCase $Parameter.Key)"
    }
    elseif ($Parameter.Value.GetType() -eq [String])
    {
        $value = "`"`"$($Parameter.Value)`"`""
    }
    elseif ($Parameter.Value.GetType() -eq [securestring])
    {
        $password = $([PSCredential]::new(0, $Parameter.Value).GetNetworkCredential().Password)
        
        $value = "`"`"$($password)`"`""
    }
    elseif ($Parameter.Value.GetType().BaseType -eq [Enum])
    {
        $value = "$(ConvertFrom-PascalCase $Parameter.Value)"
    }
    elseif ($Parameter.Value.GetType() -eq [Uri])
    {
        $value = "(POSIX file `"`"$($Parameter.Value)`"`" )"
    }
    elseif ($Parameter.Value.GetType().BaseType -eq [array])
    {
        $list = @()

        foreach ($item in $Parameter.Value)
        {
            if ($item.GetType() -eq [string])
            {
                $list += "`"`"$item`"`""
            }
            elseif ($item.GetType() -contains @([int], [Double]))
            {
                $list += $item
            }
            else 
            {
                $list += "`"$item`""
            }
        }

        $value = "{ " + $($list -join ", ") + " }"
    }
    else 
    {
        $value = $Parameter.Value
    }
    
    return " $value "
}

function New-AppleScriptCommand 
{
    <#
    
    .SYNOPSIS
    Generate an AppleScript command from a PowerShell function.

    #>
    
    param (
        [string]$Command,
        [hashtable]$Parameters,
        [array]$IgnoreParameters
    )

    if ($IgnoreParameters.Count -gt 0)
    {
        foreach ($ignoreParameter in $IgnoreParameters)
        {
            $Parameters.Remove($ignoreParameter)
        }
    }

    if ($Parameters.ContainsKey('DirectParameter'))
    {
        $DirectParameter = $Parameters.('DirectParameter')

        $value = ConvertTo-AppleScriptParameter $([Collections.DictionaryEntry]::new("DirectParameter", $DirectParameter))

        $Command += $value

        $Parameters.Remove('DirectParameter')
    }

    foreach ($parameter in $Parameters.GetEnumerator())
    {
        $value = ConvertTo-AppleScriptParameter $parameter

        if ($parameter.Value.GetType() -eq [switch])
        {
            $Command += $value 
        }
        else 
        {
            $Command += " $(ConvertFrom-PascalCase $parameter.Key) $value "
        } 
    }

    Write-Debug $Command

    $Command
}

Set-Alias -Name Invoke-AppleScript     -Value Invoke-OSA  
Set-Alias -Name Invoke-AppleScriptObjC -Value Invoke-OSA  

Export-ModuleMember -Alias @(
    'Invoke-AppleScript',
    'Invoke-AppleScriptObjC'
    )

Export-ModuleMember -Function @(
    'ConvertFrom-PascalCase',
    'ConvertTo-AppleScriptParameter',
    'ConvertTo-Hashtable',
    'ConvertTo-PascalCase',
    'ConvertTo-PSCutomObject'
    'Invoke-JavaScript'
    'Invoke-OSA',
    'New-AppleScriptCommand'
    )
