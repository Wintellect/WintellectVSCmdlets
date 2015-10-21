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
# MIIO0QYJKoZIhvcNAQcCoIIOwjCCDr4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUYDGitaQqV1RtuBnt2RPpztTP
# VjSgggmnMIIEkzCCA3ugAwIBAgIQR4qO+1nh2D8M4ULSoocHvjANBgkqhkiG9w0B
# AQUFADCBlTELMAkGA1UEBhMCVVMxCzAJBgNVBAgTAlVUMRcwFQYDVQQHEw5TYWx0
# IExha2UgQ2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMSEwHwYD
# VQQLExhodHRwOi8vd3d3LnVzZXJ0cnVzdC5jb20xHTAbBgNVBAMTFFVUTi1VU0VS
# Rmlyc3QtT2JqZWN0MB4XDTEwMDUxMDAwMDAwMFoXDTE1MDUxMDIzNTk1OVowfjEL
# MAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UE
# BxMHU2FsZm9yZDEaMBgGA1UEChMRQ09NT0RPIENBIExpbWl0ZWQxJDAiBgNVBAMT
# G0NPTU9ETyBUaW1lIFN0YW1waW5nIFNpZ25lcjCCASIwDQYJKoZIhvcNAQEBBQAD
# ggEPADCCAQoCggEBALw1oDZwIoERw7KDudMoxjbNJWupe7Ic9ptRnO819O0Ijl44
# CPh3PApC4PNw3KPXyvVMC8//IpwKfmjWCaIqhHumnbSpwTPi7x8XSMo6zUbmxap3
# veN3mvpHU0AoWUOT8aSB6u+AtU+nCM66brzKdgyXZFmGJLs9gpCoVbGS06CnBayf
# UyUIEEeZzZjeaOW0UHijrwHMWUNY5HZufqzH4p4fT7BHLcgMo0kngHWMuwaRZQ+Q
# m/S60YHIXGrsFOklCb8jFvSVRkBAIbuDlv2GH3rIDRCOovgZB1h/n703AmDypOmd
# RD8wBeSncJlRmugX8VXKsmGJZUanavJYRn6qoAcCAwEAAaOB9DCB8TAfBgNVHSME
# GDAWgBTa7WR0FJwUPKvdmam9WyhNizzJ2DAdBgNVHQ4EFgQULi2wCkRK04fAAgfO
# l31QYiD9D4MwDgYDVR0PAQH/BAQDAgbAMAwGA1UdEwEB/wQCMAAwFgYDVR0lAQH/
# BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4YxaHR0cDovL2NybC51c2Vy
# dHJ1c3QuY29tL1VUTi1VU0VSRmlyc3QtT2JqZWN0LmNybDA1BggrBgEFBQcBAQQp
# MCcwJQYIKwYBBQUHMAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZI
# hvcNAQEFBQADggEBAMj7Y/gLdXUsOvHyE6cttqManK0BB9M0jnfgwm6uAl1IT6TS
# IbY2/So1Q3xr34CHCxXwdjIAtM61Z6QvLyAbnFSegz8fXxSVYoIPIkEiH3Cz8/dC
# 3mxRzUv4IaybO4yx5eYoj84qivmqUk2MW3e6TVpY27tqBMxSHp3iKDcOu+cOkcf4
# 2/GBmOvNN7MOq2XTYuw6pXbrE6g1k8kuCgHswOjMPX626+LB7NMUkoJmh1Dc/VCX
# rLNKdnMGxIYROrNfQwRSb+qz0HQ2TMrxG3mEN3BjrXS5qg7zmLCGCOvb4B+MEPI5
# ZJuuTwoskopPGLWR5Y0ak18frvGm8C6X0NL2KzwwggUMMIID9KADAgECAhA/+9To
# TVeBHv2GK8w5hdxbMA0GCSqGSIb3DQEBBQUAMIGVMQswCQYDVQQGEwJVUzELMAkG
# A1UECBMCVVQxFzAVBgNVBAcTDlNhbHQgTGFrZSBDaXR5MR4wHAYDVQQKExVUaGUg
# VVNFUlRSVVNUIE5ldHdvcmsxITAfBgNVBAsTGGh0dHA6Ly93d3cudXNlcnRydXN0
# LmNvbTEdMBsGA1UEAxMUVVROLVVTRVJGaXJzdC1PYmplY3QwHhcNMTAxMTE3MDAw
# MDAwWhcNMTMxMTE2MjM1OTU5WjCBnTELMAkGA1UEBhMCVVMxDjAMBgNVBBEMBTM3
# OTMyMQswCQYDVQQIDAJUTjESMBAGA1UEBwwJS25veHZpbGxlMRIwEAYDVQQJDAlT
# dWl0ZSAzMDIxHzAdBgNVBAkMFjEwMjA3IFRlY2hub2xvZ3kgRHJpdmUxEzARBgNV
# BAoMCldpbnRlbGxlY3QxEzARBgNVBAMMCldpbnRlbGxlY3QwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCkXroYjDClgcwb0IBbzJNPgxvmbD9p/y3KsFml
# OCUaSufECEh0nKtVqN+3sfdlXytYuBxZP4lDsEbwfp1ppBfeemIiXWDh0ZQYEJYq
# u3/YWqrYNyMJKeeJz7KRvN8pV4N2u+nAIDPVJFfjSqA17ZYRVZs8FigRDgcYJpnA
# GkBDjIWTKkBwc/Nhk9w1XKhDFfZwvvnYeCnNZkvPxslEOu/5p5WWJW0nWpvT9BY/
# b9PR/JDRsdnFrlvZuzrk7NDyNvDMczKCUzSnHHZh60ttRV13Raq0gDaKsSrcPk6p
# AN/HsPJQAUQNBWP+3BWmV6YFfQbCfKmZZBF4Sf/q5SdXsDA7AgMBAAGjggFMMIIB
# SDAfBgNVHSMEGDAWgBTa7WR0FJwUPKvdmam9WyhNizzJ2DAdBgNVHQ4EFgQU5qYw
# jjsOnxQvFZoWoZfp6sy4XuIwDgYDVR0PAQH/BAQDAgeAMAwGA1UdEwEB/wQCMAAw
# EwYDVR0lBAwwCgYIKwYBBQUHAwMwEQYJYIZIAYb4QgEBBAQDAgQQMEYGA1UdIAQ/
# MD0wOwYMKwYBBAGyMQECAQMCMCswKQYIKwYBBQUHAgEWHWh0dHBzOi8vc2VjdXJl
# LmNvbW9kby5uZXQvQ1BTMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmwudXNl
# cnRydXN0LmNvbS9VVE4tVVNFUkZpcnN0LU9iamVjdC5jcmwwNAYIKwYBBQUHAQEE
# KDAmMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5jb21vZG9jYS5jb20wDQYJKoZI
# hvcNAQEFBQADggEBAEh+3Rs/AOr/Ie/qbnzLg94yHfIipxX7z1OCyhW7YNTOCs48
# EIXXzaJxvQD57O+S3HoHB2ZGA1cZokli6oAQNnLeP51kxQJKcTVyL2sSkKSV/2ev
# YtImhRTRCZMXe0OrGdL3Ry7x9EaaiRrhwfVJBGbqeeWc6cprFGkkDm7KpKKoCxjv
# DF3fkQ1V0QEJXQLTnEndQB+cLKIlP+swWQQxYLhfg+P8tQ+qwAbnBNYZ7+L5TiwZ
# 8Pp0S6+T94SiuoG85E1oaQUtNT1SO8FLQa4M3bO5xdGA2GL1Vti/W8Gp8tIPr/wM
# Ak4Xt++emsid5THDZkjSrFMqbCHmaxoTmtcutr4xggSUMIIEkAIBATCBqjCBlTEL
# MAkGA1UEBhMCVVMxCzAJBgNVBAgTAlVUMRcwFQYDVQQHEw5TYWx0IExha2UgQ2l0
# eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMSEwHwYDVQQLExhodHRw
# Oi8vd3d3LnVzZXJ0cnVzdC5jb20xHTAbBgNVBAMTFFVUTi1VU0VSRmlyc3QtT2Jq
# ZWN0AhA/+9ToTVeBHv2GK8w5hdxbMAkGBSsOAwIaBQCgeDAYBgorBgEEAYI3AgEM
# MQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQB
# gjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqGSIb3DQEJBDEWBBQLA7pvjr6Vzsao
# oXKWbf1Rbn2c0TANBgkqhkiG9w0BAQEFAASCAQCDmMcCeNAqhnyojXjcvOx6i8pm
# 6PfmuSwtvp4SpKpAE+HMIh1OytdWzAvWaCMzbKxtERD8EQpFUHLV9a08OQNvFssW
# BeyW8/4PbqqhmmYP6oEqu5Od5y2E3m1tlzaULWAcgosPe1NQSuAm8uaeN5HizAnk
# 97PsMqbmfjHXt/N0Hg9cnESWYl8M5uetaUf16IXefHeoz18may7246vkW8QPCN0O
# /zFD2foikVw0c/WAGfxRb6zfB4b9LRR3qNrqbgxGZwYG3qxlVHYAeFlNCv/zBI01
# yoV4AER58F/wZmkAH1mvSjB4P5cmaHiOA3T0qgkOilj8gp7BdaIVprVarWqtoYIC
# RDCCAkAGCSqGSIb3DQEJBjGCAjEwggItAgEAMIGqMIGVMQswCQYDVQQGEwJVUzEL
# MAkGA1UECBMCVVQxFzAVBgNVBAcTDlNhbHQgTGFrZSBDaXR5MR4wHAYDVQQKExVU
# aGUgVVNFUlRSVVNUIE5ldHdvcmsxITAfBgNVBAsTGGh0dHA6Ly93d3cudXNlcnRy
# dXN0LmNvbTEdMBsGA1UEAxMUVVROLVVTRVJGaXJzdC1PYmplY3QCEEeKjvtZ4dg/
# DOFC0qKHB74wCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEw
# HAYJKoZIhvcNAQkFMQ8XDTEyMDcwNjIzMTk0M1owIwYJKoZIhvcNAQkEMRYEFEcp
# 8Oan5CavYWsQxrJOEPg7hUxAMA0GCSqGSIb3DQEBAQUABIIBAF9mqqoxZpKoM8g9
# Vn4jwq/RAWsOjkEGawwYH+ZekYazQmLIQTJpDVZtTYxwT/EM2DkICRgiXxL9qA8r
# 56Enla/QZceIEbPY6UnOdwuaW67jDOj+b+ycBXoVQMu816weTp2PwtUEuzuurLxB
# gtal5fT/evnTruoHUWLn2/Qrs3KP8qf89Hx4hI8AR3j9FS33T4syH2L8+s1SQExa
# upFvDGhTfCE38JqtiazqSbcbzOs/PwbHJev9sQs5KCEjB7hapVWu9K6z9JjrEwIP
# Wp2AdcY3z99TMLb/kxkqfHQpwKF3DvjkxVA4cCLmhXSb/ZF26v1FB1zj5Y5fGj5f
# wApPGrI=
# SIG # End signature block
