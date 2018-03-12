begin {
    $component = $MyInvocation.MyCommand.Name

    # Dot source PowerShell Toolkit
    $scriptPath = "."
    if ($PSScriptRoot) {
        $scriptPath = $PSScriptRoot
    } else {
        $scriptPath = (Get-Item -Path ".\").FullName
    }

    # Dot source Aetna PowerShell Toolkit
    . $scriptPath\PowerShell-Toolkit.ps1

    # Load Global Variables
    Start-GlobalVariables

    # Script Name
    $scriptFileName = ([io.fileinfo]$MyInvocation.MyCommand.Definition).BaseName

    # Initiate Logging Environment
    $Global:logFileName = "$scriptFileName.log"
    $Global:logDirectory = "$env:SystemDrive\zco\inst"
    $Global:logFile = Join-Path $logDirectory $logFileName
    $Global:scriptName = "$scriptFileName.ps1"

    Start-Script
   
}

Process {

}

end {
    Finish-Script -ExitCode $Global:exitCode
}