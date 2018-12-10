#requires -version 2.0
###############################################################################
# Wintellect NuGet Cmdlets Module
#
# The macro replacements for Dev 11
#
# Copyright (c) 2002 - 2012 - John Robbins/Wintellect
# 
# Do whatever you want with this module, but please do give credit.
###############################################################################


# Always make sure all variables are defined and all best practices are 
# followed.
Set-StrictMode -version Latest

###############################################################################
function Get-AppPools
{
<#
.SYNOPSIS
Returns the names of all IIS application pools.
#>

  cd C:\windows\system32\inetsrv
  
  .\appcmd list apppool | foreach { $_ | Select-String -pattern '".*"' } | foreach { $_.Matches.Value } | Sort-Object
}

Export-ModuleMember -function Get-AppPools
###############################################################################

###############################################################################
function Get-Threads
{
<#
.SYNOPSIS
Returns the threads for the current process
.DESCRIPTION
All threads for the current process being debugged are returned. If the 
debugger is not active, returns $null
#>

    # Check if we are debugging. 2 == dbgDebugMode.dbgBreakMode
    if ($dte.Debugger.CurrentMode -ne 2)
    {
        Write-Warning "Get-Threads only works when debugging."
        return $null
    }
    
    $dte.Debugger.CurrentProgram.Threads
}

Export-ModuleMember -function Get-Threads
###############################################################################

###############################################################################
function Debug-IISProcess
{
<#
.SYNOPSIS
Attaches Visual Studio to the w3wp.exe process associated with a given application pool name.
#>

    param (
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$AppPoolName
    ) 

    $process = Get-WmiObject -Namespace "root/cimv2" -Query "SELECT * FROM Win32_Process where Name = 'w3wp.exe'" |  Where-Object { $_.CommandLine -match "-ap `"$AppPoolName`"" } 

    if ($process -eq $null) {
        Write-Host "Could not find w3wp.exe process associated with app pool '$AppPoolName'."
    } else {
        Write-Host "Attaching to w3wp.exe process associated with app pool '$AppPoolName'."
        
        foreach ($proc in $dte.Debugger.LocalProcesses) {
            if ($proc.ProcessID -eq $process.ProcessId){
                $proc.Attach()
                break
            }
        }    
    }
}

Export-ModuleMember -function Debug-IISProcess
###############################################################################

###############################################################################
function Get-Breakpoints
{
<#
.SYNOPSIS
Returns all breakpoints
.DESCRIPTION
The default $dte.Debugger.Breakpoints returns the older version of the 
breakpoint interface. This cmdlet promotes all breakpoints to the the 
most recent interface to allow access to properties such as FilterBy.
#>

    # The magic here is in the wonderful NuGet Get-Interface cmdlet.
    # There's no way this integration would work without it because
    # of issues with PowerShell.
    $dte.Debugger.BreakPoints | `
        ForEach-Object { Get-Interface $_ ([ENVDTE80.Breakpoint2]) }
}

Export-ModuleMember -function Get-Breakpoints 
###############################################################################

###############################################################################
function Disable-NonActiveThreads
{
<#
.SYNOPSIS
Freezes all threads except the active thread.
.DESCRIPTION
When multithreaded debugging it's common to be single stepping and bounce to 
another thread. The Disable-NonActiveThreads, and it's counterpart, 
Resume-NonActiveThreads, make multithreaded debugging easier by freezing and
thawing all other threads, except for the active thread. This way you can 
finish single stepping through a method and avoid the bounce.
.LINK
http://www.wintellect.com/CS/blogs/jrobbins/archive/2009/07/17/automatically-freezing-threads-brrrrr.aspx
#>

    # Check if we are debugging. 2 == dbgDebugMode.dbgBreakMode
    if ($dte.Debugger.CurrentMode -ne 2)
    {
        Write-Warning "Disable-NonActiveThreads only works when debugging."
        return
    }
    
    $currThread = $dte.Debugger.CurrentThread.Id
    
    Get-Threads | Where-Object { $_.ID -ne $currThread } | `
            ForEach-Object { $_.Freeze() }
}

Export-ModuleMember -function Disable-NonActiveThreads
###############################################################################

###############################################################################
function Resume-NonActiveThreads
{
<#
.SYNOPSIS
Freezes all threads except the active thread.
.DESCRIPTION
When multithreaded debugging it's common to be single stepping and bounce to 
another thread. The Resume-NonActiveThreads, and it's counterpart, 
Disable-NonActiveThreads, make multithreaded debugging easier by freezing and
thawing all other threads, except for the active thread. This way you can 
finish single stepping through a method and avoid the bounce.
.LINK
http://www.wintellect.com/CS/blogs/jrobbins/archive/2009/07/17/automatically-freezing-threads-brrrrr.aspx
#>

    # Check if we are debugging. 2 == dbgDebugMode.dbgBreakMode
    if ($dte.Debugger.CurrentMode -ne 2)
    {
        Write-Warning "Resume-NonActiveThreads only works when debugging."
        return
    }
    
    $currThread = $dte.Debugger.CurrentThread.Id
    
    Get-Threads | Where-Object { $_.ID -ne $currThread } | `
            ForEach-Object { $_.Thaw() }
}

Export-ModuleMember -function Resume-NonActiveThreads 
###############################################################################

###############################################################################
$script:k_THREADFILTER = "ThreadName == InterestingThread"

function Add-InterestingThreadFilterToBreakpoints
{
<#
.SYNOPSIS
Adds the breakpoint filter "ThreadName==InterestingThread" to all breakpoints.
.DESCRIPTION
Multithreading debugging is hard enough. Having a bunch of breakpoints set 
means you are going stop no matter which thread is executing those locations.
You know which thread is the interesting thread so this cmdlet, and it's 
opposite, Remove-InterestingThreadFilterFromBreakpoints, will add and remove
a filter "ThreadName==InterestingThread" to all breakpoints except those that
already have a filter. To name your thread, go into the VS Threads window,
right click on the thread, and rename it to InterestingThread. 
.LINK
http://www.wintellect.com/CS/blogs/jrobbins/archive/2009/07/12/easier-multithreaded-debugging.aspx
#>
    Get-Breakpoints | `
        Where-Object {$_.FilterBy.Length -eq 0} | `
            ForEach-Object {$_.FilterBy = $k_THREADFILTER}
}

Export-ModuleMember -function Add-InterestingThreadFilterToBreakpoints 
###############################################################################

###############################################################################
function Remove-InterestingThreadFilterFromBreakpoints
{
<#
.SYNOPSIS
Removes the breakpoint filter "ThreadName==InterestingThread" to all breakpoints.
.DESCRIPTION
Multithreading debugging is hard enough. Having a bunch of breakpoints set 
means you are going stop no matter which thread is executing those locations.
You know which thread is the interesting thread so this cmdlet, and it's 
opposite, Add-InterestingThreadFilterFromBreakpoints, will add and remove
a filter "ThreadName==InterestingThread" to all breakpoints except those that
already have a filter. To name your thread, go into the VS Threads window,
right click on the thread, and rename it to InterestingThread. 
.LINK
http://www.wintellect.com/CS/blogs/jrobbins/archive/2009/07/12/easier-multithreaded-debugging.aspx
#>
    Get-Breakpoints | `
        Where-Object {$_.FilterBy -eq $k_THREADFILTER} | `
            ForEach-Object {$_.FilterBy = ""}
}

Export-ModuleMember -function Remove-InterestingThreadFilterFromBreakpoints
###############################################################################
$script:k_DebuggerCommands = ".loadby sos clr;.load sosex;.echo **TO DETACH USE THE qd COMMAND**"

function Invoke-WinDBG
{
<#
.SYNOPSIS
Attaches WinDBG to the current process being debugged.
.DESCRIPTION
While using Visual Studio, there are times you need to use WinDBG, such as to
look at memory with SOS or SOSEX. This cmdlet will attach WinDBG to the current
process through a non-invasive attach. Once you are finished using the excellent
informational commands in WinDBG, use QD to go back to debugging in VS.
.PARAMETER WinDBG 
The default is to assume WinDBG is in the path, but if it is not, specify it
in this parameter. Also, use this parameter to specifically use the x86
version if the x64 version is in the path.
.LINK
http://www.wintellect.com/cs/blogs/jrobbins/archive/2007/12/08/use-two-debuggers-at-once.aspx 
#>
    param
    (
        [Parameter(HelpMessage="The optional full path to WINDBG.EXE")]
        [string] $WinDBG = "WinDBG.EXE"
    )
    
    # Check if we are debugging. 2 == dbgDebugMode.dbgBreakMode
    if ($dte.Debugger.CurrentMode -ne 2)
    {
        Write-Warning "Invoke-WinDBG only works when debugging or when stopped in the debugger."
        return
    }
    
    # Get the current process as an IProcess2.
    $proc = Get-Interface $dte.Debugger.CurrentProcess ([ENVDTE80.Process2])
    
    # If the TrasportQualifier is not equal to the machine name, the process
    # is being remote debugged so I can't put WinDBG on it.
    if ($proc.TransportQualifier -ne $ENV:COMPUTERNAME)
    {
        Write-Warning "Invoke-WinDBG only works doing local debugging."
        Write-Warning "The current process is running remotely."
        return
    }
    
    $cmdLine = '-pv -p {0} -c "{1}"' -f $proc.ProcessID, $k_DebuggerCommands
   
    Start-Process $WinDBG -ArgumentList $cmdLine
}

Export-ModuleMember -function Invoke-WinDBG
###############################################################################

###############################################################################
Function Invoke-NamedParameter 
{
<#
.SYNOPSIS
Call methods with optional parameters easily in PowerShell
.DESCRIPTION
COM methods with many optional parameters are hard to call in PowerShell, this
awesome cmdlet from Jason Archer makes is very simple. See the full discussion
at the link.
.EXAMPLE
Calling a method with named parameters.

$shell = New-Object -ComObject Shell.Application
Invoke-NamedParameter $Shell "Explore" @{"vDir"="$pwd"}

## the syntax for more than one would be @{"First"="foo";"Second"="bar"}

.EXAMPLE
Calling a method that takes no parameters (you can also use -Argument with $null).

$shell = New-Object -ComObject Shell.Application
Invoke-NamedParameter $Shell "MinimizeAll" @{}
.LINK
http://stackoverflow.com/questions/5544844/how-to-call-a-complex-com-method-from-powershell
#>
    [CmdletBinding(DefaultParameterSetName = "Named")]
    param(
        [Parameter(ParameterSetName = "Named", Position = 0, Mandatory = $true)]
        [Parameter(ParameterSetName = "Positional", Position = 0, Mandatory = $true)]
        [ValidateNotNull()]
        [System.Object]$Object
        ,
        [Parameter(ParameterSetName = "Named", Position = 1, Mandatory = $true)]
        [Parameter(ParameterSetName = "Positional", Position = 1, Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [String]$Method
        ,
        [Parameter(ParameterSetName = "Named", Position = 2, Mandatory = $true)]
        [ValidateNotNull()]
        [Hashtable]$Parameter
        ,
        [Parameter(ParameterSetName = "Positional")]
        [Object[]]$Argument
    )

    end {  ## Just being explicit that this does not support pipelines
        if ($PSCmdlet.ParameterSetName -eq "Named") {
            ## Invoke method with parameter names
            ## Note: It is ok to use a hashtable here because the keys (parameter names) and values (args)
            ## will be output in the same order.  We don't need to worry about the order so long as
            ## all parameters have names
            $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod,
                $null,  ## Binder
                $Object,  ## Target
                ([Object[]]($Parameter.Values)),  ## Args
                $null,  ## Modifiers
                $null,  ## Culture
                ([String[]]($Parameter.Keys))  ## NamedParameters
            )
        } else {
            ## Invoke method without parameter names
            $Object.GetType().InvokeMember($Method, [System.Reflection.BindingFlags]::InvokeMethod,
                $null,  ## Binder
                $Object,  ## Target
                $Argument,  ## Args
                $null,  ## Modifiers
                $null,  ## Culture
                $null  ## NamedParameters
            )
        }
    }
}

Export-ModuleMember -function Invoke-NamedParameter
###############################################################################

###############################################################################
function Script:RecurseCodeElements($elems, $bpHash , $srcFile, $tagValue)
{    
    foreach ($currElem in $elems)
    {
        $currElem = Get-Interface $currElem ([EnvDTE80.CodeElement2])
        
        # vsCMElement.vsCMElementClass = 1
        # vsCMElement.vsCMElementNamespace = 5
        # vsCMElement.vsCMElement = 11
        if (1,5,11 -contains $currElem.Kind)
        {
            $subCodeElems = $null
            try
            {
                $subCodeElems = $currElem.Children
            }
            catch
            {
                $subCodeElems = $null
            }
            
            if ($subCodeElems -ne $null)
            {
                RecurseCodeElements $subCodeElems $bpHash $srcFile $tagValue
            }
        }
        # vsCMElement.vsCMElementFunction = 2
        # vsCMElement.vsCMElementProperty = 4
        elseif (2,4 -contains $currElem.Kind)
        {
            # Attributed COM component attributes show up pulled out into 
            # their functions. The only thing is that their StartPoint 
            # property is invalid and throws an exception when accessed.
            $txtPoint = $null
            
            try
            {
                $txtPoint = $currElem.StartPoint
            }
            catch
            {
                $txtPoint = $null
            }
            
            if ($txtPoint -ne $null)
            {
                # Check to see if there's a breakpoint already set as VS
                # is not happy setting breakpoints on top of existing 
                # breakpoints.
                $line = $txtPoint.Line.ToString()
                if ($bpHash["$srcFile + $line"] -eq $null)
                {
                    try
                    {
                        $bps = Invoke-NamedParameter $dte.Debugger.Breakpoints "Add" @{"File"="$srcFile";"Line"="$line"}
                        
                        # Now set the tag to identify the breakpoint for clearing it.
                        foreach($bp in $bps)
                        {
                            $bp.Tag = $tagValue
                        }
                    }
                    catch
                    {
                        # If there's an error setting a BP, it's because the
                        # BP was attempted on a DllImport attribute item so
                        # ignore it.
                    }
                }
            }
        }
    }
}

function script:GetActiveDocument
{
    $activeDoc = $dte.ActiveDocument
    if ($activeDoc -eq $null)
    {
        Write-Warning "There's no active document window."
        return $null
    }
    
    if ($activeDoc.ProjectItem.FileCodeModel -eq $null)
    {
        Write-Warning "There is no code model associated with the current document."
        return $null
    }
    $activeDoc
}

function Add-BreakpointsOnAllDocMethods
{
<#
.SYNOPSIS
Sets a breakpoint on each method in the current code document.
.DESCRIPTION
While native code supports setting breakpoints on all methods in a code 
document, .NET languages do not. The cmdlet takes care of that for you. The
idea is that you can use the debugger to ensure you're calling all methods
in a class when debugging/testing. Using the debugger for function coverage
is easy with this cmdlet.
#>
    $activeDoc = GetActiveDocument
    if ($activeDoc -eq $null)
    {
        return
    }
    
    $srcFile = $activeDoc.FullName 
    $tagValue = "Wintellect Rocks! " + $srcFile 
    
    # I'm going to have to look at the breakpoint list a lot so let's
    # hash that bad boy to make it faster.
    $bpHash = @{}
    
    Get-Breakpoints | ForEach-Object { $bpHash[$_.File + $_.FileLine.ToString()] = $_.FileLine }
    
    $fileCodeModel2 = Get-Interface $activeDoc.ProjectItem.FileCodeModel ([ENVDTE80.FileCodeModel2])
    $codeElements = $fileCodeModel2.CodeElements
    RecurseCodeElements $codeElements $bpHash $srcFile $tagValue
}

Export-ModuleMember -function Add-BreakpointsOnAllDocMethods
###############################################################################

###############################################################################
function Remove-BreakpointsOnAllDocMethods
{
<#
.SYNOPSIS
Removes breakpoints previously set with Add-BreakpointsOnAllMethods
#>

    # Since I used the tag property on the breakpoint I can just look for
    # those breakpoints and clear them.
    
    $activeDoc = GetActiveDocument
    if ($activeDoc -eq $null)
    {
        return
    }
    
    $srcFile = $activeDoc.FullName 
    $tagValue = "Wintellect Rocks! " + $srcFile 
    
    Get-Breakpoints | Where-Object { $_.Tag -eq $tagValue } | ForEach-Object { $_.Delete() }
}

Export-ModuleMember -function Remove-BreakpointsOnAllDocMethods 
###############################################################################

###############################################################################
# Another Jason Archer special! :)
# http://stackoverflow.com/questions/5648931/powershell-test-if-registry-value-exists
Function Script:Test-RegistryValue {
    param(
        [Alias("PSPath")]
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
        [String]$Path
        ,
        [Parameter(Position = 1, Mandatory = $true)]
        [String]$Name
        ,
        [Switch]$PassThru
    ) 

    process {
        if (Test-Path $Path) {
            $Key = Get-Item -LiteralPath $Path
            if ($Key.GetValue($Name, $null) -ne $null) {
                if ($PassThru) {
                    Get-ItemProperty $Path $Name
                } else {
                    $true
                }
            } else {
                $false
            }
        } else {
            $false
        }
    }
}

function Open-LastIntelliTraceRecording
{
<#
.SYNOPSIS
Opens the last IntelliTrace log
.DESCRIPTION
IntelliTrace is great, but when you stop debugging the current debugging 
session's IntelliTrace log goes poof and disappears. This is very 
frustrating because the time you really want to look at an IntelliTrace
log is after you finish debugging. This cmdlet opens the last run 
IntelliTrace log.

In order for this cmdlet to work, you need to go to set an option to
enable saving the files otherwise they are gone when the debugging session
ends.

Go to Tools/Options/IntelliTrace/Advanced and check 
'Store IntelliTrace recordings in this directory'

Note that turning this setting on in Dev 11 Beta brings up a bug where the
storage directory is not properly cleaned up when Dev 11 ends. Keep an eye
on the directory as the IntelliTrace files get quite large.
.LINK
http://www.wintellect.com/CS/blogs/jrobbins/archive/2009/10/19/vs-2010-beta-2-intellitrace-in-depth-first-look.aspx
#>

    # Check if we are in design mode. 1 == dbgDebugMode.dbgDesign
    if ($dte.Debugger.CurrentMode -ne 1)
    {
        Write-Warning "Open-LastIntelliTraceRecording only works when not debugging."
        return $null
    }

    $dir = $null

    $ver = $dte.Version
    if ($ver -lt "14.0")
    {
        $regPath = "HKCU:\Software\Microsoft\VisualStudio\$ver\DialogPage\Microsoft.VisualStudio.TraceLogPackage.ToolsOptionAdvanced"
        if ( !(Test-Path $regPath ) -or
            ( (Get-ItemProperty -Path $regPath)."SaveRecordings" -eq "False") )
        {
            Write-Warning "You must set IntelliTrace to save the recordings."
            Write-Warning "Go to Tools/Options/IntelliTrace/Advanced and check 'Store IntelliTrace recordings in this directory'"
            return 
        }
    
        if ( !(Test-RegistryValue $regPath "RecordingPath") )
        {
            Write-Warning "The RecordingPath property does not exist or is not set."
            return
        }
    
        $dir = (Get-ItemProperty -Path $regPath)."RecordingPath"
    }
    else
    {
        $regPath = "HKCU:\SOFTWARE\Microsoft\VisualStudio\$ver\ApplicationPrivateSettings\Microsoft\VisualStudio\TraceLogPackage\ToolsOptionAdvanced"
        if ( !(Test-Path $regPath) -or
              ((Get-ItemProperty -Path $regPath)."SaveRecordings" -match "False"))
        {
            Write-Warning "You must set IntelliTrace to save the recordings."
            Write-Warning "Go to Tools/Options/IntelliTrace/Advanced and check 'Store IntelliTrace recordings in this directory'"
            return 
        }

        if ( !(Test-RegistryValue $regPath "RecordingPath") )
        {
            Write-Warning "The RecordingPath property does not exist or is not set."
            return
        }
    
        $dir = ((Get-ItemProperty -Path $regPath)."RecordingPath").Substring("1*System.String*".Length)     
    }

    # Get all the filenames from the recording path.
    $fileNames = Get-ChildItem -Path $dir | Sort-Object LastWriteTime -Descending
    if ($fileNames -ne $null)
    {
        
        # If the user has VSHOST debugging turned on for WPF/Console/WF apps, 
        # current instance will be sitting there with no access set. I'll try opening
        # in order until I get one to open. This accounts for multiple instances 
        # of VS running as they all share the same directory.
        for ($i = 0 ; $i -lt $fileNames.Length; $i++)
        {
            $toOpen = $fileNames[$i].FullName
            try
            {
                [void]$dte.ItemOperations.OpenFile($toOpen)
                return
            }
            catch 
            {
            }
        }
    }
    else
    {
        Write-Warning "No IntelliTrace files are present"
    }
}

Export-ModuleMember -function Open-LastIntelliTraceRecording
###############################################################################

# SIG # Begin signature block
# MIIYTQYJKoZIhvcNAQcCoIIYPjCCGDoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUMcqXbgdO9kEFseclB0oE0VKo
# u9qgghM9MIIEhDCCA2ygAwIBAgIQQhrylAmEGR9SCkvGJCanSzANBgkqhkiG9w0B
# AQUFADBvMQswCQYDVQQGEwJTRTEUMBIGA1UEChMLQWRkVHJ1c3QgQUIxJjAkBgNV
# BAsTHUFkZFRydXN0IEV4dGVybmFsIFRUUCBOZXR3b3JrMSIwIAYDVQQDExlBZGRU
# cnVzdCBFeHRlcm5hbCBDQSBSb290MB4XDTA1MDYwNzA4MDkxMFoXDTIwMDUzMDEw
# NDgzOFowgZUxCzAJBgNVBAYTAlVTMQswCQYDVQQIEwJVVDEXMBUGA1UEBxMOU2Fs
# dCBMYWtlIENpdHkxHjAcBgNVBAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEhMB8G
# A1UECxMYaHR0cDovL3d3dy51c2VydHJ1c3QuY29tMR0wGwYDVQQDExRVVE4tVVNF
# UkZpcnN0LU9iamVjdDCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAM6q
# gT+jo2F4qjEAVZURnicPHxzfOpuCaDDASmEd8S8O+r5596Uj71VRloTN2+O5bj4x
# 2AogZ8f02b+U60cEPgLOKqJdhwQJ9jCdGIqXsqoc/EHSoTbL+z2RuufZcDX65OeQ
# w5ujm9M89RKZd7G3CeBo5hy485RjiGpq/gt2yb70IuRnuasaXnfBhQfdDWy/7gbH
# d2pBnqcP1/vulBe3/IW+pKvEHDHd17bR5PDv3xaPslKT16HUiaEHLr/hARJCHhrh
# 2JU022R5KP+6LhHC5ehbkkj7RwvCbNqtMoNB86XlQXD9ZZBt+vpRxPm9lisZBCzT
# bafc8H9vg2XiaquHhnUCAwEAAaOB9DCB8TAfBgNVHSMEGDAWgBStvZh6NLQm9/rE
# JlTvA73gJMtUGjAdBgNVHQ4EFgQU2u1kdBScFDyr3ZmpvVsoTYs8ydgwDgYDVR0P
# AQH/BAQDAgEGMA8GA1UdEwEB/wQFMAMBAf8wEQYDVR0gBAowCDAGBgRVHSAAMEQG
# A1UdHwQ9MDswOaA3oDWGM2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9BZGRUcnVz
# dEV4dGVybmFsQ0FSb290LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUHMAGG
# GWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEFBQADggEBAE1C
# L6bBiusHgJBYRoz4GTlmKjxaLG3P1NmHVY15CxKIe0CP1cf4S41VFmOtt1fcOyu9
# 08FPHgOHS0Sb4+JARSbzJkkraoTxVHrUQtr802q7Zn7Knurpu9wHx8OSToM8gUmf
# ktUyCepJLqERcZo20sVOaLbLDhslFq9s3l122B9ysZMmhhfbGN6vRenf+5ivFBjt
# pF72iZRF8FUESt3/J90GSkD2tLzx5A+ZArv9XQ4uKMG+O18aP5cQhLwWPtijnGMd
# ZstcX9o+8w8KCTUi29vAPwD55g1dZ9H9oB4DK9lA977Mh2ZUgKajuPUZYtXSJrGY
# Ju6ay0SnRVqBlRUa9VEwggSUMIIDfKADAgECAhEAn+rIEbDxYkel/CDYBSOs5jAN
# BgkqhkiG9w0BAQUFADCBlTELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlVUMRcwFQYD
# VQQHEw5TYWx0IExha2UgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3
# b3JrMSEwHwYDVQQLExhodHRwOi8vd3d3LnVzZXJ0cnVzdC5jb20xHTAbBgNVBAMT
# FFVUTi1VU0VSRmlyc3QtT2JqZWN0MB4XDTE1MDUwNTAwMDAwMFoXDTE1MTIzMTIz
# NTk1OVowfjELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3Rl
# cjEQMA4GA1UEBxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQx
# JDAiBgNVBAMTG0NPTU9ETyBUaW1lIFN0YW1waW5nIFNpZ25lcjCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBALw1oDZwIoERw7KDudMoxjbNJWupe7Ic9ptR
# nO819O0Ijl44CPh3PApC4PNw3KPXyvVMC8//IpwKfmjWCaIqhHumnbSpwTPi7x8X
# SMo6zUbmxap3veN3mvpHU0AoWUOT8aSB6u+AtU+nCM66brzKdgyXZFmGJLs9gpCo
# VbGS06CnBayfUyUIEEeZzZjeaOW0UHijrwHMWUNY5HZufqzH4p4fT7BHLcgMo0kn
# gHWMuwaRZQ+Qm/S60YHIXGrsFOklCb8jFvSVRkBAIbuDlv2GH3rIDRCOovgZB1h/
# n703AmDypOmdRD8wBeSncJlRmugX8VXKsmGJZUanavJYRn6qoAcCAwEAAaOB9DCB
# 8TAfBgNVHSMEGDAWgBTa7WR0FJwUPKvdmam9WyhNizzJ2DAdBgNVHQ4EFgQULi2w
# CkRK04fAAgfOl31QYiD9D4MwDgYDVR0PAQH/BAQDAgbAMAwGA1UdEwEB/wQCMAAw
# FgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDov
# L2NybC51c2VydHJ1c3QuY29tL1VUTi1VU0VSRmlyc3QtT2JqZWN0LmNybDA1Bggr
# BgEFBQcBAQQpMCcwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5j
# b20wDQYJKoZIhvcNAQEFBQADggEBAA27rWARG7XwDczmSDp6Pg4z3By56tYg/qNN
# 0Mx2TugY2Hnf00+aQmQjiilyijpsZqY8OheocEVlxnPD0M6JVPusaQ9YsBnLhp9+
# uX7rUZK/m93r0WXwJXuIfN69pci1FFG8wIEwioU4e+Z5/mdVk4f+T+iNDu3zcpK1
# womAbdFZ4x0N6rE47gOdABmlqyGbecPMwj5ofr3JTWlNtGRR+7IodOJTic6d+q3i
# 286re34GRHT9CqPJt6cwzUnSkmTxIqa4KEV0eemnzjsz+YNQlH1owB1Jx2B4ejxk
# JtW++gpt5B7hCVOPqcUjrMedYUIh8CwWcUk7EK8sbxrmMfEU/WwwggTnMIIDz6AD
# AgECAhAQcJ1P9VQI1zBgAdjqkXW7MA0GCSqGSIb3DQEBBQUAMIGVMQswCQYDVQQG
# EwJVUzELMAkGA1UECBMCVVQxFzAVBgNVBAcTDlNhbHQgTGFrZSBDaXR5MR4wHAYD
# VQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsxITAfBgNVBAsTGGh0dHA6Ly93d3cu
# dXNlcnRydXN0LmNvbTEdMBsGA1UEAxMUVVROLVVTRVJGaXJzdC1PYmplY3QwHhcN
# MTEwODI0MDAwMDAwWhcNMjAwNTMwMTA0ODM4WjB7MQswCQYDVQQGEwJHQjEbMBkG
# A1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAwDgYDVQQHEwdTYWxmb3JkMRowGAYD
# VQQKExFDT01PRE8gQ0EgTGltaXRlZDEhMB8GA1UEAxMYQ09NT0RPIENvZGUgU2ln
# bmluZyBDQSAyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAy/jnp+jx
# lyhAaIA30sg/jpKKkjeHR4DqTJnPbvkVR73udfRErNDD1E33GcDTPE3BR7lZZRaT
# jNkKhJuf6PZqY1j+X9zRf0tRnwAcAIdUIAdXoILJL5ivM4q7e4AiJWpsr8IsbHkT
# vaMqSNa1jmFV6WvoPYC/FAOFGI5+TOnCGYhzknLN+v9QTcsspnsac7EAkCzZMuL7
# /ayVQjbsNMUTU2iywZ9An9p7yJ1ibJOiQtd5n5dPMVtQIaGrr9kcss51vlssVgAk
# jRHBdR/w/tKV/vDhMSMYZ8BbE/1amJSU//9ZAh8ArObx8vo6c7MdQvxUdc9RMS/j
# 24HZdyMqT1nOIwIDAQABo4IBSjCCAUYwHwYDVR0jBBgwFoAU2u1kdBScFDyr3Zmp
# vVsoTYs8ydgwHQYDVR0OBBYEFB7FsSx9h9oCaHwlvAwHhD+2z97xMA4GA1UdDwEB
# /wQEAwIBBjASBgNVHRMBAf8ECDAGAQH/AgEAMBMGA1UdJQQMMAoGCCsGAQUFBwMD
# MBEGA1UdIAQKMAgwBgYEVR0gADBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8vY3Js
# LnVzZXJ0cnVzdC5jb20vVVROLVVTRVJGaXJzdC1PYmplY3QuY3JsMHQGCCsGAQUF
# BwEBBGgwZjA9BggrBgEFBQcwAoYxaHR0cDovL2NydC51c2VydHJ1c3QuY29tL1VU
# TkFkZFRydXN0T2JqZWN0X0NBLmNydDAlBggrBgEFBQcwAYYZaHR0cDovL29jc3Au
# dXNlcnRydXN0LmNvbTANBgkqhkiG9w0BAQUFAAOCAQEAlYl3k2gBXnzZLTcHkF1a
# Ql4MZLQ2tQ/2q9U5J94iRqRJHGZLRhlZLnlJA/ackt9tUDVcDJEuYANZ0PFk92kJ
# 9n7+6zSzbbG/ZpyjujF4uYc1YT2SMRvv9Oie1qxF+gw2PIBnu73vLsKQ4T1xLzvB
# sFh+RcNScQMH9vM5TYs2IRsB39naXivrDpeAHkQcUIj1xhIzSqhNpY0vlAx7xr+a
# LMMyzb2MJybw4TADUAaCvPQ7s4N1Bsbvuu7TgPhSxqzLefI4nnuwklhCkQXIliGt
# uUsWgRRp8Tew/jT33LDfl/VDEJt2j7Rl9eifE7cerG/EaYpfujxhfl5JhiMTLq8V
# SDCCBS4wggQWoAMCAQICEHF/qKkhW4DS4HFGfg8Z8PIwDQYJKoZIhvcNAQEFBQAw
# ezELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4G
# A1UEBxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxITAfBgNV
# BAMTGENPTU9ETyBDb2RlIFNpZ25pbmcgQ0EgMjAeFw0xMzEwMjgwMDAwMDBaFw0x
# ODEwMjgyMzU5NTlaMIGdMQswCQYDVQQGEwJVUzEOMAwGA1UEEQwFMzc5MzIxCzAJ
# BgNVBAgMAlROMRIwEAYDVQQHDAlLbm94dmlsbGUxEjAQBgNVBAkMCVN1aXRlIDMw
# MjEfMB0GA1UECQwWMTAyMDcgVGVjaG5vbG9neSBEcml2ZTETMBEGA1UECgwKV2lu
# dGVsbGVjdDETMBEGA1UEAwwKV2ludGVsbGVjdDCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBAMFQoSYu2olPhQGXgsuq0HBwHsQBoFbuAoYfX3WVp2w8dvji
# kqS+486CmTx2EMH/eKbgarVP0nGIA266BNQ5GXxziGKGk5Y+g74dB269i8G2B24X
# WXZQcw0NTch6oUcXuq2kOkcp1srh4Pp+HQB/qR33qQWzEW7yMlpoI+SwNoa9p1WQ
# aOPzoAfJdiSgInWGgrlAxVwcET0AmVQQKQ2lgJyzQkXIAiRxyJPSgKbZrhTa7/BM
# m33SWmG9K5GlFaw76HFV1e49v8hrTDFJJ7CAQz65IcazjqHTaKOfYhsPhiFrm/Ap
# kPUuJb45MeEPms8DzD8lTSQfo7eLkG2hNtxkRmcCAwEAAaOCAYkwggGFMB8GA1Ud
# IwQYMBaAFB7FsSx9h9oCaHwlvAwHhD+2z97xMB0GA1UdDgQWBBQEi+PkyNipSO6M
# 0oxTXEhobEPaWzAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADATBgNVHSUE
# DDAKBggrBgEFBQcDAzARBglghkgBhvhCAQEEBAMCBBAwRgYDVR0gBD8wPTA7Bgwr
# BgEEAbIxAQIBAwIwKzApBggrBgEFBQcCARYdaHR0cHM6Ly9zZWN1cmUuY29tb2Rv
# Lm5ldC9DUFMwQQYDVR0fBDowODA2oDSgMoYwaHR0cDovL2NybC5jb21vZG9jYS5j
# b20vQ09NT0RPQ29kZVNpZ25pbmdDQTIuY3JsMHIGCCsGAQUFBwEBBGYwZDA8Bggr
# BgEFBQcwAoYwaHR0cDovL2NydC5jb21vZG9jYS5jb20vQ09NT0RPQ29kZVNpZ25p
# bmdDQTIuY3J0MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20w
# DQYJKoZIhvcNAQEFBQADggEBAB4m8FXuYk3D2mVYZ3vvghqRSRVgEqJmG7YBBv2e
# QCk9CYML37ubpYigH3JDzWMIDS8sfv6hzJzY4tgRuY29rJBMyaWRw228IEOLkLZq
# Si/JOxOT4NOyLYacOSD1DHH63YFnlDFpt+ZRAOKbPavG7muW97FZT3ebCvLCJYrL
# lYSym4E8H/y7ICSijbaIBt/zHtFX8RJvV7bijvxZI1xqqKyx9hyF/4gNWMq9uQiE
# wIG13VT/UmNCc3KcCsy9fqnWreFh76EuI9arj1VROG2FaYQdaxD2O+9nl+uxFmOM
# eOHqhQWlv57eO9do7PI6PiVGMTkiC2eFTeBEHWylCUFDkDIxggR6MIIEdgIBATCB
# jzB7MQswCQYDVQQGEwJHQjEbMBkGA1UECBMSR3JlYXRlciBNYW5jaGVzdGVyMRAw
# DgYDVQQHEwdTYWxmb3JkMRowGAYDVQQKExFDT01PRE8gQ0EgTGltaXRlZDEhMB8G
# A1UEAxMYQ09NT0RPIENvZGUgU2lnbmluZyBDQSAyAhBxf6ipIVuA0uBxRn4PGfDy
# MAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3
# DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEV
# MCMGCSqGSIb3DQEJBDEWBBRg70J01tKKfgvLqQGcRbDK4PODYTANBgkqhkiG9w0B
# AQEFAASCAQCK6t4zK+tEXEGy3F9zMzAGGCA9EjVRVvDYiE58sQQ/2HFRIUX66pUR
# Nw7el46r0bQbTkbaZ0K22qgUkkR1jh8mIluLLaHBI0BFV5K8L6WxXRhnISVaHNzk
# M9rhiEEteieWQ7iZkt9E2th2z8zTCH1i2eMBZEedU+IjUy+HkZY5aTzrl3LdyWt+
# C1G9tgVDG0p6BA51jjOMz/zs6gI5iV0gjzwnv8opqDXeTa1OSYv+pK7StQZeNCFH
# swWPmzsafCCFstf+cDWi3hIGnH0VzEyCubaKweeWGW4qY6H0O3B+GTCA7i9RNu0v
# avkUZNkl4BYzsvioaXkO/jlF3HVnPY29oYICRTCCAkEGCSqGSIb3DQEJBjGCAjIw
# ggIuAgEAMIGrMIGVMQswCQYDVQQGEwJVUzELMAkGA1UECBMCVVQxFzAVBgNVBAcT
# DlNhbHQgTGFrZSBDaXR5MR4wHAYDVQQKExVUaGUgVVNFUlRSVVNUIE5ldHdvcmsx
# ITAfBgNVBAsTGGh0dHA6Ly93d3cudXNlcnRydXN0LmNvbTEdMBsGA1UEAxMUVVRO
# LVVTRVJGaXJzdC1PYmplY3QCEQCf6sgRsPFiR6X8INgFI6zmMAkGBSsOAwIaBQCg
# XTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0xNTEw
# MjEyMzI2NDJaMCMGCSqGSIb3DQEJBDEWBBSHo4OHNdrCxPqlI9h292PgDLl4ijAN
# BgkqhkiG9w0BAQEFAASCAQAjThGOsnqK2WKB2xr8C/gRMIrfHb2ifjDguDkm5DHD
# Fyi8dX/0oe0pLrC5SR7fhLsIfPsjR9XpA7pY6Ssb4iT/Od97R4X1roAXMcb6nmYe
# qmNVr/qhaF/2eJsIkB8QbZV3DBorhnftS74/BXHD9qHSk3Z4qZsxxP/2tJQ9Eyh2
# FgwGLU5BrDAX9FwrpQ2gLloXAHqLyNP3Q3/Uci1VzF+Gjjd6l7DstpRZUnLi96tx
# Ofbu7Nb2Q4tu2mFEB1LxzALTOFdwRhs7fByQ3b/Mf//fkY62s9z1DNvMRzkgcCYR
# 2uCBBvTBW3YVGiun1Ma4SLYT4y4dolsIwJWp36vOnWYu
# SIG # End signature block
