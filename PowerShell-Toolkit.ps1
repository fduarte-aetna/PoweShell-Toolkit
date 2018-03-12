Function Start-Script {
    
    $component = $MyInvocation.MyCommand.Name
   
    Write-Log -Message "======================================================" -Component $component -type 1
    Write-Log -Message "                 ---Starts Install--                     " -Component $component -type 1
    Write-Log -Message "======================================================" -Component $component -type 1

    $Global:exitCode = 0

}


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


Function Write-Log {
    [CmdletBinding()]
    Param(
        [parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [String]
        $Message,

        [parameter()]
        [String]
        $Component = $MyInvocation.MyCommand.Name,

        [parameter()]
        [int]
        $Type = 1
    )

    Begin {
        if (!(Test-Path $Global:logDirectory)) {
            try {
                New-Item -Path  $Global:logDirectory -ItemType Directory -Force -ErrorAction Stop
            } catch {
                $eventLogParam = @{
                    LogName = "Application"
                    Source = "PowerShell-Toolkit"
                    EventId = "1111"
                    Message = "Unable to write create directory $Global:logDirectory. Exception was: $($Error[0].Exception)"
                    EntryType = "Error"}
                Write-EventLog @eventLogParam -Verbose
            }
        }
    }

    Process {
        Write-Verbose "Start Process"

        try {

            Write-Verbose "Start Try"
            # Build time variables
            [string]$time = Get-Date -Format "HH:mm:ss.fff"

            # Build time zone data
            [string]$tzb = (Get-WmiObject -Query "Select Bias from Win32_TimeZone").Bias


            Write-Verbose "Building Write-Log message"
            # Build string to log
            [string]$toLog = "<![LOG[{0}]LOG]!><time=`"{1}`" date=`"{2}`" component=`"{3}`" context=`"`" type=`"{4}`" thread=`"{5}`" file=`"`">" -f ($Message), ("$time+$($tzb.SubString(1,3))"), (Get-Date -Format "MM-dd-yyyy"), ($Component),($type),($PID)

            
            Write-Verbose "Writing to log file"
            # Log string to log file
            $toLog | Out-File -Encoding default -Append -NoClobber -FilePath ("filesystem::{0}" -f $Global:LogFile) -ErrorAction Stop

            Write-Verbose "Stop Try"
        } catch {
            Write-Verbose "Start Catch"

            if([Environment]::UserInteractive) {
                Write-Error -Exception $Error[0].Exception
            } else {
                $eventLogParam = @{
                    LogName = "Application"
                    Source = "Aetna-PowerShell-Toolkit"
                    EventId = "1111"
                    Message = "Unable to write to log file $($Global:LogFile). Exception was: $($Error[0].Exception)"
                    EntryType = "Error"
                }
                Write-Verbose "Logging to event log: $($Error[0].Exception)"
                Write-EventLog @eventLogParam
            }

            Write-Verbose "Stop Catch"
        }

        Write-Verbose "Stop Process"
    }

    End {
        Write-Verbose "Start End"
        Write-Verbose "Stop End"
    }
}


Function Execute-Process {
    [CmdletBinding()]
    Param (
        [parameter(Mandatory=$true,
        ValueFromPipeline=$true,
        ValueFromPipelineByPropertyName=$true)]
        [String]
        $Path,

        [parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [String]
        $Parameters,
        
        [parameter(Mandatory=$false)]
        [AllowEmptyString()]
        [string]
        $WorkingDirectory
    )

    begin {
        Write-Verbose "Begin"
        
        $component = $MyInvocation.MyCommand.Name

        if (!(Test-Path $Path)) {
            $Global:exitCode = 2 # The system cannot find the file specified.

            Write-Verbose "Error: $exitCode File path does not exist: `"$Path`""
            Write-Log -Message "File path does not exist: `"$Path`"" -Component $component -Type 3
            
            exit($exitCode)
        }
    } process {
        
        Write-Verbose "Process"

        try {
            Write-Verbose "Start Try"

            $startExec = New-Object System.Diagnostics.ProcessStartInfo
            $startExec.FileName = $Path
            $startExec.Arguments = $Parameters
            $startExec.WorkingDirectory = $WorkingDirectory
            $startExec.UseShellExecute = $false

            Write-Log -Message "Executing: `"$Path`"" -Component $component -Type 1
            Write-Log -Message "Parameters: `"$Parameters`"" -Component $component -Type 1
            Write-Log -Message "Working Directory: `"$WorkingDirectory`"" -Component $component -Type 1

            Write-Verbose "Executing: `"$Path`""
            Write-Verbose "Parameters: `"$Parameters`""
            Write-Verbose "Working Directory: `"$WorkingDirectory`""

            $exec = New-Object System.Diagnostics.Process
            $exec.StartInfo = $startExec
            $exec.Start() | Out-Null
            $exec.WaitForExit()

            $Global:exitCode = $exec.ExitCode

            Write-Verbose "Stop Try"

        
        } catch {
            Write-Verbose "Exception Caught"
            Write-Verbose "Message: $($Error[0].Exception)"
            Write-Log -Message "Unable to execute process." -component $component -type 3
            Write-Log -Message "Error message: $($Error[0].Exception)" -component $component -type 3
        }
    } end {
        Write-Verbose "End"
        Write-Verbose "Exit Code: $exitCode"
        if ($exitCode -ne 0) {
            Write-Log -Message "$component did not exit with a 0 exit code" -Component $component -Type 3
            Write-Log -Message "Exit code was $exitCode" -Component $component -Type 3
        } else {
            Write-Log -Message "Exit code was $exitCode" -Component $component -Type 1
        }
    }
}


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

Function Finish-Script {
    [CmdletBinding()]
    Param(
        [parameter()]
        [int]
        $ExitCode
    )

    $component = $MyInvocation.MyCommand.Name

    if ($ExitCode -ne 0) {
        Write-Log -Message "Exit code: $ExitCode" -Component $component -Type 3
    } else {
        Write-Log -Message "Exit code: $ExitCode" -Component $component -Type 1
    }

    Write-Log -Message "--------------------End Install--------------------------" -Component $component -Type 1

    Exit $ExitCode

}