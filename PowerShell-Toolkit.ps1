<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Start-GlobalVariables {

    # Initiate Shell Applicaton
    $shellApp = New-Object -ComObject Shell.Application

    # Get environment variables
    $procArch6432 = [System.Environment]::ExpandEnvironmentVariables("%PROCESSOR_ARCHITEW6432%")
    
    if ($procArch6432.ToUpper() -eq 'AMD64') {
        # 32-bit process in a 64-bit OS
        $Global:osArch = 64
        $Global:procBitness = 32
        $Global:isProcess6432 = $true
        $Global:programFiles32 = ($shellApp.Namespace(0x2a)).Self.Path
        $Global:commonProgFiles32 = ($shellApp.Namespace(0x2c)).Self.Path
        $Global:systemDir32 = ($shellApp.Namespace(0x29)).Self.Path
    } else {
        $procArch = [System.Environment]::ExpandEnvironmentVariables("%PROCESSOR_ARCHITECTURE%")
        
        if (($procArch.ToUpper() -eq 'X86') -or ($procArch.ToUpper() -eq 'ARM')) {
            $Global:osArch = 32
            $Global:procBitness = 32
            $Global:isProcess6432 = $false
            $Global:programFiles32 = ($shellApp.Namespace(0x2a)).Self.Path
            $Global:commonProgFiles32 = ($shellApp.Namespace(0x2c)).Self.Path
            $Global:systemDir32 = ($shellApp.Namespace(0x29)).Self.Path
        } elseif ($procArch.ToUpper() -eq 'AMD64') {
            $Global:osArch = 64
            $Global:procBitness = 64
            $Global:isProcess6432 = $false
            $Global:programFiles32 = ($shellApp.Namespace(0x2a)).Self.Path
            $Global:programFiles64 = ($shellApp.Namespace(0x26)).Self.Path
            $Global:commonProgFiles32 = ($shellApp.Namespace(0x2c)).Self.Path
            $Global:commonProgFiles64 = ($shellApp.Namespace(0x2b)).Self.Path
            $Global:systemDir32 = ($shellApp.Namespace(0x29)).Self.Path
            $Global:systemDir64 = ($shellApp.Namespace(0x25)).Self.Path
        }
    }

    # Computer Name
    $Global:computerName = $env:COMPUTERNAME

    # User Name
    $Global:userName = $env:USERNAME
    
    # %ProgramData%\Microsoft\Windows\Start Menu
    $Global:allUsersStartMenu = ($shellApp.Namespace(0x16)).Self.Path

    # %ProgramData%\Microsoft\Windows\Start Menu\Programs
    $Global:allUsersStartMenuPrograms = ($shellApp.Namespace(0x17)).Self.Path

    # %SystemDrive%\Users\Public\Desktop
    $Global:allUsersDesktop = ($shellApp.Namespace(0x19)).Self.Path
    
    # %SystemDrive%\ProgramData
    $Global:programData = ($shellApp.Namespace(0x23)).Self.Path
    
    # %WINDIR%
    $Global:windowsFolder = ($shellApp.Namespace(0x24)).Self.Path
}

<#
.SYNOPSIS
    Short description

.DESCRIPTION
    Long description

.PARAMETER Message
    Parameter description

.PARAMETER Component
    Parameter description

.PARAMETER Type
    Parameter description

.EXAMPLE
    An example

.NOTES
    General notes
#>
Function Write-Log {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [String]$Message,

        [parameter()]
        [String]$Component = $MyInvocation.MyCommand.Name,

        [parameter()]
        [int]$Type = 1
    )

    Begin {}

    Process {
        Write-Debug "Starting Write-Log Process"
        Write-Verbose "Starting Process"

        try {

            Write-Debug "Building Write-Log variables"
            Write-Verbose "Building Write-Log variables"
            # Build time variables
            [string]$time = Get-Date -Format "HH:mm:ss.fff"

            # Build time zone data
            [string]$tzb = (Get-WmiObject -Query "Select Bias from Win32_TimeZone").Bias


            Write-Debug "Building Write-Log message"
            Write-Verbose "Building Write-Log message"
            # Build string to log
            [string]$toLog = "<![LOG[{0}]LOG]!><time=`"{1}`" date=`"{2}`" component=`"{3}`" context=`"`" type=`"{4}`" thread=`"{5}`" file=`"`">" -f ($Message), ("$time+$($tzb.SubString(1,3))"), (Get-Date -Format "MM-dd-yyyy"), ($Component),($type),($PID)

            
            Write-Debug "Writing to log file"
            Write-Verbose "Writing to log file"
            # Log string to log file
            $toLog | Out-File -Encoding default -Append -NoClobber -FilePath ("filesystem::{0}" -f $Global:LogFile) -ErrorAction Stop
        } catch {
            if([Environment]::UserInteractive) {
                Write-Error -Exception $Error[0].Exception
            } else {
                Write-Debug "Session isn't interactive."
                $eventLogParam = @{
                    LogName = "Application"
                    Source = "Aetna-PowerShell-Toolkit"
                    EventId = "1111"
                    Message = "Unable to write to log file $($Global:LogFile). Exception was: $($Error[0].Exception)"
                    EntryType = "Error"
                }
                                
                Write-EventLog @eventLogParam
            }

        }

        Write-Debug "Finishing Write-Log Process"
        Write-Verbose "Finishing Write-Log Process"
    }

    End {}

}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER FilePath
Parameter description

.PARAMETER Parameters
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Start-Operation {
    Param (
        [Parameter()]
        [string]$FilePath,
        [Parameter()]
        [string]$Parameters
    )

    $component = $MyInvocation.MyCommand.Name

    try {
        $startExec = New-Object System.Diagnostics.ProcessStartInfo
        $startExec.FileName = $FilePath
        $startExec.Arguments = $Parameters
        $startExec.CreateNoWindow = $false
        $startExec.UseShellExecute = $false

        Write-Log -Message "Executing: `"$FilePath`" $Parameters" -Component $component -Type 1

        $exec = New-Object System.Diagnostics.Process
        $exec.StartInfo = $startExec
        $exec.Start() | Out-Null
        $exec.WaitForExit()

        $Global:exitCode =$exec.ExitCode

        Write-Log -Message "Execution exit code: $exitCode" -Component $component -Type 1
        
    } catch {
        Write-Log -Message "Unable to start process. Error message was: $($Error[0].Exception)" -component $component -type 3
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Source
Parameter description

.PARAMETER Target
Parameter description

.PARAMETER Overwrite
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Copy-File {
    Param (
        [Parameter()]
        [string]$Source,
        [Parameter()]
        [string]$Target,
        [Parameter()]
        [switch]$Overwrite

    )

    $component = $MyInvocation.MyCommand.Name

    if ($Source -notlike "filesystem::*") {
        $Source = "filesystem::$Source"
    }

    if ($Target -notlike "filesystem::*") {
        $Target = "filesystem::$Target"
    }

    $TargetFolder = $Target.Replace($Target.Split("\")[-1],"")

    try {
        if (!(Test-Path -Path $TargetFolder)) {
            New-Item $TargetFolder -ItemType Directory -ErrorAction Stop
        }
    } catch {
        Write-Log -Message "Unable to create directory: $TargetFolder" -component $component -type 3
        Write-Log -Message "Error Message: $($Error[0].Exception)" -component $component -type 3
    }

    try {
        if ($Overwrite) {
            Copy-Item -Path $Source -Destination $Target -Recurse -Force -ErrorAction Stop
        } else {
            Copy-Item -Path $Source -Destination $Target -ErrorAction Stop
        }
    } catch {
        Write-Log -Message "Unable to copy source: $Source to $Target Error Message: $($Error[0].Exception)" -component $component -type 3
    }  
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Set-Permissions {
    param (
        [Parameter()]
        $Path
    )

    $component = $MyInvocation.MyCommand.Name
    
    try {
        $acl = Get-Acl $Path
        $ar = New-Object System.Security.AccessControl.FileSystemAccessRule("Users","FullControl","ContainerInherit,ObjectInherit","None","Allow")
        $acl.SetAccessRule($ar)
        Set-Acl $Path $acl -ErrorAction Stop
    } catch {
        Write-Log -Message "Error Message: $(Error[0].Exception)" -component $component -type 3
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Remove-File {
    Param (
        [Parameter()]
        $Path
    )

    $component = $MyInvocation.MyCommand.Name

    try {
        if (Test-Path -Path $Path -PathType Leaf) {
            Write-Log -Message "Found file: $Path and will now be delete it." -component $component -type 1
            Remove-Item -Path $Path -Force -ErrorAction Stop
        } else {
            Write-Log -Message "File path: $Path was not found on the syste" -component $component -type 1
        }
    } catch {
        Write-Log -Message "Unable to delete $Path because of the following error: $($Error[0].Exception)" -component $component -type 3
    }
}


<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Remove-Directory {
    Param (
        [Parameter()]
        $Path
    )

    $component = $MyInvocation.MyCommand.Name

    try {
        if (Test-Path -Path $Path -PathType Container) {
            Write-Log -Message "Found Folder Path: $Path. Will now recursively delete it." -component $component -type 1
            Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
        } else {
            Write-Log -Message "Folder path: $Path was not found on the system." -component $component -type 1
        }
    } catch {
        Write-Log -Message "Unable to delete $Path because of the following error: $($Error[0].Exception)" -component $component -type 3
    }
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER Path
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Test-FileLocked {
    Param (
        [string]$Path
    )

    try {
        [IO.File]::OpenWrite($Path).Close()
        return $false

    } catch {
        return $true
    }
}


<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER ProductGUID
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Get-MSI {
    Param (
        [Parameter()]
        $ProductGUID
    )

    $component = $MyInvocation.MyCommand.Name

    $PathNative = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$ProductGUID"
    $PathRedirected = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\$ProductGUID"
    
    if (Test-Path $PathNative) {
        Write-Log "Found: $PathNative" -Component $component -Type 1
        return $true
    } elseif (Test-Path $PathRedirected) {
        Write-Log "Found: $PathRedirected" -Component $component -Type 1
        return $true
    } else {
        Write-Log "MSI Prodcut Code: $ProductGUID not found" -Component $component -Type 1
        return $false
    }
}