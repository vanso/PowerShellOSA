# PowerShellOSA

**PowerShellOSA** is a PowerShell module.

This module enables seamless **interoperability between the Open Scripting Architecture (OSA) and PowerShell on macOS**. It allows users to execute OSA language scripts (such as AppleScript) or script files directly from PowerShell, as well as invoke PowerShell commands from within an OSA language. Communication between the two scripting environments is handled transparently, enabling smooth automation workflows and data exchange. This makes it easier to combine macOS-native automation capabilities with the flexibility and power of PowerShell scripting.

## Requirements

- PowerShell 7.2 or later

## Supported Platforms

- macOS 10.15 or later (Intel and Apple Silicon)

## Installation

```powershell
PS> Install-Module -Name PowerShellOSA (coming soon)
```

## Available Functions in the Module

| Function name            | Description                                                                          |
|:-------------------------|:-------------------------------------------------------------------------------------|
| ConvertTo-AppleScriptParameter | Convert a PowerShell parameter to an AppleScript parameter. |
| Invoke-JavaScript              | Specific usage of Invoke-OSA in JavaScript. |
| Invoke-OSA | Allowing PowerShell to call OSA script or file using the PowerShellOSA application through seamless two-way communication. |
| New-AppleScriptCommand | Generate an AppleScript command from a PowerShell function. |

## Available Aliases in the Module

| Alias name | Function name  |                                                                
|:-------------------------|:-------------------------------------------------------------------------------------|
|  Invoke-AppleScript           | Invoke-OSA     |
|  Invoke-AppleScriptObjC       | Invoke-OSA  |

## Supported Types

| AppleScript           | PowerShell    |
|:-------------------------|:-------------|
| missing value |$null| 
| boolean |[Boolean]| 
| text |[String]| 
| integer | [[SByte, Byte, UInt16, Int16, UInt32, Int32, UInt64, Int64, UInt128, UInt128, UIntPtr, IntPtr, Numerics.BigInteger]](https://learn.microsoft.com/en-us/dotnet/standard/numerics#integer-types) | 
| real  | [[Single, Double, Decimal]](https://learn.microsoft.com/en-us/dotnet/standard/numerics#floating-point-types). The Half type isn't supported. | 
| date | [DateTime] |
| file | [Uri] (Input) / [String] (Output)|
| alias | [String] (Output)|
| list | [Array] | 
| record | [Hashtable] |

## Using AppleScript from PowerShell

```powershell
Invoke-OSA [[-Script] <String>] [-Path <String>] [-InputObject <Object>] [-Parameters <Object>] [-Language <String>] 
[-OutputFormat {PSCustomObject | JSON | PLIST}] [<CommonParameters>]
```
A set of functions for using AppleScript features from PowerShell is provided by the [PSMacToolkit](https://github.com/vanso/PSMacToolkit) module.

### Usage (AppleScript & JavaScript for Automation)

```powershell
Invoke-AppleScript "current date"

Invoke-AppleScript "system info" -OutputFormat PLIST

Invoke-AppleScript @'
    tell application "Image Events"
        launch
        get properties of display profile of every display
    end tell
'@

Invoke-AppleScript @'
tell application "Contacts"
	get name of every person whose birth date is not missing value
end tell
'@

Invoke-JavaScript "var app = Application.currentApplication()
            app.includeStandardAdditions = true
            
            app.displayDialog('What\'s your name?', {
                defaultAnswer: app.systemInfo().longUserName,
                withIcon: 'note',
                buttons: ['Cancel', 'Continue'],
                defaultButton: 'Continue' })"
```

### Using an Explicitly Defined `run` Handler with Parameters in AppleScript

Parameters of the run handler are passed as a list, and each element can have any arbitrary name (even a single character).  
Example: `run(a, b)`  

**input : object**  
This parameter contains the objects that are passed either through the pipeline or explicitly via the InputObject parameter.

**parameters : object**  
This parameter contains the parameters of your choice, which can be nested.

```powershell
Invoke-AppleScript "on run (input, parameters)
        display dialog last item of (DayNames of DateTimeFormat of input) as text
    end run" -InputObject $(Get-Culture)

Invoke-AppleScript "on run (input, parameters)
        display dialog `"`"Temporary items path: `"`" & (|temporary items path| of parameters) as text
    end run" -Parameters $(@{ "temporary items path" = [Uri]::new([System.IO.Path]::GetTempFileName()).ToString() })

Invoke-AppleScript "on run (input, parameters)
        display dialog `"`"Hello `"`" & (item 1 of parameters) as text
    end run" -Parameters @("Sal", "Sue", "Yoshi", "Wayne", "Carla") })

Invoke-AppleScript "on run (input, parameters)
        display dialog `"`"Current date: `"`" & (parameters) as text
    end run" -Parameters (Get-Date)
```

### Using an Explicitly Defined `run` Handler with Parameters in JavaScript

```powershell
Invoke-JavaScript "function run(input, parameters)
{
	 // See above for more information.
}"

Invoke-OSA "function run(input, parameters)
{
    // See above for more information.
	
}" -Language JavaScript
```

### Using AppleScriptObjC

```powershell
Invoke-AppleScriptObjC @'
	use framework "Foundation"
	
	set theString to "Hello World"
	set theString to stringWithString_(theString) of NSString of current application
	set theString to (uppercaseString() of theString) as string
	return theString
'@
```
## Using PowerShell from AppleScript

### AppleScript Dictionary

<img width="2142" height="2198" alt="dico" src="https://github.com/user-attachments/assets/44d77975-0471-4b4f-93fa-2560c6d581d5" />

### Supported Output Streams

| **Stream** #             | **Description** | **Write Cmdlet** | **Supported** |
|:------------------------:|:-------------|:--------|:----:|
| 1	  | Success stream	   | 	Write-Output | ✅ | 
| 2   | Error stream	   | 	Write-Error | ✅ | 
| 3	  | Warning stream	   | 	Write-Warning | ✅ | 
| 4	  | Verbose stream	   | 	Write-Verbose | ✅ | 
| 5	  | Debug stream	   | 	Write-Debug | ✅ | 
| 6	  | Information stream | 	Write-Information | ❌ | 
| n/a | Progress stream    | 	Write-Progress | ⚠️ (The progress bar is limited to one level.)| 

### User Interaction Utilities

| **Command** | **Supported** |
|:--------|:----:|
| Write-Host | ⚠️ | 
| Read-Host | ⚠️ | 
| Get-Credential | ✅ | 
| Confirm | ❌ | 
| ShouldProcess | ❌ | 
| ShouldContinue | ❌ | 
| PromptForChoice | ✅ | 

Use the functions provided by the [PSMacToolkit](https://github.com/vanso/PSMacToolkit) PowerShell module for user interaction instead.

✅ Full support<br />
⚠️ Partial support<br />
❌ Not supported

### Usage

```applescript
tell application "PowerShellOSA"
	
	do pwsh script "Param(
 					   [int]$Count = 10  # Number of Fibonacci terms to calculate (starting from F0)
					)

					# --- Define Constants for Binet's Formula ---
					$sqrt5 = [Math]::Sqrt(5)
					$phi = (1 + $sqrt5) / 2 # The Golden Ratio (Phi)

					# --- Calculate and Display the Sequence ---
					for ($n = 0; $n -lt $Count; $n++) 
					{
  						# Calculate the approximation using the simplified Binet's Formula: Phi^n / sqrt(5)
    						$approximation = ([Math]::Pow($phi, $n) / $sqrt5)
    
    						# Round the approximation to the nearest integer and cast to a 64-bit integer ([int64])
    						Write-Output $([int64][Math]::Round($approximation))
					}" with parameters 42
end tell
```
```applescript
tell application "PowerShellOSA"
	
	do pwsh script "[CmdletBinding()]
					Param( 
                    	[Parameter(ValueFromPipeline)]
						[int]$Number
					)
		
   					BEGIN
    				{
						Write-Output \"In Begin block\"
   	 				}
 
					PROCESS
   	 				{
    	    			Write-Output \"In Process block\"
					}
					END
					{
						Write-Output \"In End block\"
  	  				}
					" with input {1, 2, 3}
end tell
```

## License

This module is released under the terms of the GNU General Public License (GPL), Version 2.
