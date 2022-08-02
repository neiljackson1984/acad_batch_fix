#!pwsh

# override some of the default parameters, on Neil's machine only:
if($env:computername -match "AUTOSCAN-WK07W7"){
    $pathOfAcadExecutable = join-path $env:ProgramFiles "Autodesk\AutoCAD 2023\acad.exe"
    $pathOfAcadCoreConsoleExecutable = join-path $env:ProgramFiles "Autodesk\AutoCAD 2023\accoreconsole.exe"
    # $pathOfDirectoryInWhichToSearchForDwgFiles='J:\Nakano\2022-02-08_projectwise_aec_acad_troubleshoot\testroot'
    $pathOfDirectoryInWhichToSearchForDwgFiles='C:\work\nakano_marginal_way'
} else {
    $pathOfAcadExecutable = join-path $env:ProgramFiles "Autodesk/AutoCAD 2018/acad.exe"
    $pathOfAcadCoreConsoleExecutable = join-path $env:ProgramFiles "Autodesk/AutoCAD 2018/accoreconsole.exe"
    $pathOfDirectoryInWhichToSearchForDwgFiles='C:\bms'    
}



$pathOfWorkingDirectoryInWhichToRunAcad=$PSScriptRoot

$pathOfLogFile = "$(join-path $pathOfDirectoryInWhichToSearchForDwgFiles $MyInvocation.MyCommand.Name).log"

$skipToOrdinal = $null
# $skipToOrdinal = 237
# a hack for resuming our place in the sequence, intended for use in debugging.
# unassign or set to $null to to not skip anything

$doObscureNonactiveDwgFiles = $True





function writeToLog($x){
    $stringToAppend = "$(Get-Date -Format "yyyy/MM/dd HH:mm:ss"): $x"
    Add-Content -Path $pathOfLogFile -Value "$stringToAppend"
    Write-Host "$stringToAppend"
}


function setRegistryValueInAllAcadProfiles($relativePathAndName, $value){
    $registryPathOfUserKey  = Join-Path "HKEY_USERS" $((Get-Item -Path "registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI").GetValue("LastLoggedOnUserSID"))
    $registryPathOfAutocadKey = Join-Path $registryPathOfUserKey "Software\Autodesk\Autocad"
    # $registryPathOfProfileKeys = Join-Path $registryPathOfAutocadKey '*\*\Profiles\*\Variables'
    $registryPathOfProfileKeys = Join-Path $registryPathOfAutocadKey '*\*\Profiles\*'
    $profileKeys = @(Get-Item -path "registry::$registryPathOfProfileKeys")
    foreach ($profileKey in $profileKeys)
    {
        $registryPathOfKeyOfInterest = Join-Path $profileKey.Name (Split-Path  -Path $relativePathAndName -Parent)
        $name = Split-Path  -Path $relativePathAndName -Leaf
        writeToLog ("now setting $name to $value in $registryPathOfKeyOfInterest")
        # Get-Item -path "registry::$registryPathOfVariablesKey" | Set-ItemProperty -Name $name -Value $value -Type String
        Get-Item -path "registry::$registryPathOfKeyOfInterest" | Set-ItemProperty -Name $name -Value $value
    }
}

function getPlainGuidString(){
    (New-Guid).Guid -replace "-",""
}



writeToLog "$($MyInvocation.MyCommand.Name) is running."



$pathOfMyOwnTempDirectory = join-path $env:TEMP "fix_$(Get-Date -Format "yyyyMMdd_HHmm")_$(getPlainGuidString)"
New-Item -ItemType "directory" -Path $pathOfMyOwnTempDirectory | Out-Null

$pathOfTemporaryAcadConsoleLogDirectory    = join-path $pathOfMyOwnTempDirectory "log_accumulators/acad_console_log_accumulator"
$pathOfTemporarySlapdownLogDirectory       = join-path $pathOfMyOwnTempDirectory "log_accumulators/slapdown_log_accumulator"
$pathOfStdoutCollectorFile                 = join-path $pathOfMyOwnTempDirectory "log_accumulators/stdout_accumulator"
$pathOfStderrCollectorFile                 = join-path $pathOfMyOwnTempDirectory "log_accumulators/stderr_accumulator"
$pathOfReviewableAcadConsoleLogDirectory   = join-path $pathOfMyOwnTempDirectory "acad_console_logs"

New-Item -ItemType "directory" -Path $pathOfTemporaryAcadConsoleLogDirectory | Out-Null
New-Item -ItemType "directory" -Path $pathOfTemporarySlapdownLogDirectory    | Out-Null
New-Item -ItemType "directory" -Path $pathOfReviewableAcadConsoleLogDirectory    | Out-Null


$script:popupSlapdownScriptProcess = $null
$script:acadProcess = $null

function flushAcadConsoleLog(){
    $returnValue = ""
    foreach($logFile in (Get-ChildItem -Path $pathOfTemporaryAcadConsoleLogDirectory -File -Recurse )){
        $returnValue += (Get-Content -Raw -Path $logFile.FullName)
        Remove-Item -Path $logFile.FullName  | Out-Null
    }
    $returnValue
}

function sweepAcadConsoleLogToOurLog(){
    # foreach($logFile in (Get-ChildItem -Path $pathOfTemporaryAcadConsoleLogDirectory -File -Recurse )){
    #     writeToLog "content of $($logFile.FullName):`n$(Get-Content -Raw -Path $logFile.FullName)"
    #     Remove-Item -Path $logFile.FullName 
    # }

    # writeToLog (flushAcadConsoleLog)
    writeToLog "acad console log: `n$(indent (flushAcadConsoleLog) -doLineNumbers $true)"
}




function flushSlapdownLog(){
    $returnValue = ""
    foreach($logFile in (Get-ChildItem -Path $pathOfTemporarySlapdownLogDirectory -File -Recurse )){
        $returnValue += (Get-Content -Raw -Path $logFile.FullName)
        Remove-Item -Path $logFile.FullName | Out-Null
    }
    $returnValue
}

function sweepSlapdownLogToOurLog(){
    writeToLog "slapdown log: `n$(indent (flushSlapdownLog) -doLineNumbers $true)"
}

function flushAcadErrFile(){
    $returnValue = ""
    $pathOfAcadErrFile=Join-Path $pathOfWorkingDirectoryInWhichToRunAcad "acad.err"
    if(Test-Path -Path $pathOfAcadErrFile -PathType Leaf){
        $returnValue += (Get-Content -Raw -Path $pathOfAcadErrFile)
        Remove-Item -Path $pathOfAcadErrFile | Out-Null
    }
    $returnValue
}


function sweepAcadErrFileToOurLog(){
    writeToLog "acad err file content: `n$(indent (flushAcadErrFile) -doLineNumbers $true)"
}

function flushStdOutCollectorFile(){
    $returnValue = ""
    if(Test-Path -Path $pathOfStdoutCollectorFile -PathType Leaf){
        $returnValue += (Get-Content -Raw -Encoding unicode -Path $pathOfStdoutCollectorFile)
        Remove-Item -Path $pathOfStdoutCollectorFile | Out-Null
    }
    $returnValue
}

function flushStdErrCollectorFile(){
    $returnValue = ""
    if(Test-Path -Path $pathOfStderrCollectorFile -PathType Leaf){
        $returnValue += (Get-Content -Raw -Encoding unicode -Path $pathOfStderrCollectorFile)
        Remove-Item -Path $pathOfStderrCollectorFile | Out-Null
    }
    $returnValue
}

function sweepStandardStreamCollectorFilesToOurLog(){
    writeToLog "content of stdout collector file:`n$(indent (flushStdOutCollectorFile) -doLineNumbers $true)"
    writeToLog "content of stderr collector file:`n$(indent (flushStdErrCollectorFile) -doLineNumbers $true)"
}

function toHumanReadableDataSize($byteCount){
    '{0:N0} kilobytes' -f ($byteCount/([Math]::Pow(10,3)))
}

function toHumanReadableDuration($timeSpan){
    "$([math]::Round( $timeSpan.TotalSeconds )) seconds"
}

function splitStringIntoLines($s){
    # $startTime=Get-Date
    @($s -split "\r?\n")
    # $endTime=Get-Date
    # writeToLog "splitting took $(toHumanReadableDuration($endTime - $startTime))."
}

function indent($x, $indentLevel=1, $doLineNumbers=$false){
    # $startTime=Get-Date
    $returnValue = ""
    $tabString = "    "
    # foreach($line in (splitStringIntoLines $x)){
    #     $returnValue += ($tabString + $line + "`n")
    # }
    $lines = splitStringIntoLines $x

    if($doLineNumbers){
        $i = 1
        $lineNumberFieldWidth = "$($lines.Count)".Length
        $lineNumberFormattingString = "{0,$($lineNumberFieldWidth):}"
        $returnValue = ( @( $lines | foreach-object { $tabString + ( $lineNumberFormattingString -f ($i++) ) + ": " + $_} ) -join "`n" )       
    } else {
        $returnValue = (@( $lines | foreach-object {$tabString + $_} ) -join "`n")
    }
    # $endTime=Get-Date
    # writeToLog "indenting took $(toHumanReadableDuration($endTime - $startTime))."
    $returnValue
}

$popupSlapdownScriptContent = @"
#SingleInstance, Force
#Persistent
SetWorkingDir, %A_ScriptDir%
writeToLog("starting")


SetTitleMatchMode RegEx


; define the criteria for the various user-interface modal (i.e. requiring user intervention before resuming the program)
; windows that acad throws up to block our progress.

GroupAdd, lostSheetSetAssociationWindowDefinition ,Sheet Set Manager - Lost Set Association ahk_exe acad.exe
GroupAdd, fatalError1WindowDefinition             ,AutoCAD Error Aborting,FATAL ERROR:
GroupAdd, fatalError2WindowDefinition             ,AutoCAD Alert,AutoCAD cannot continue
GroupAdd, projectwiseSavePromptWindowDefinition   ,ProjectWise,Save changes to Drawing1.dwg
GroupAdd, drawingFileNotValidWindowDefinition     ,AutoCAD Message ahk_exe acad.exe,Drawing file is not valid.
; GroupAdd, openDrawingErrorsFoundWindowDefinition  ,Open Drawing - Errors Found,Would you like to cancel this open
GroupAdd, openDrawingErrorsFoundWindowDefinition  ,Open Drawing - Errors Found

GroupAdd, anyPopupWindowDefinition                ,ahk_group lostSheetSetAssociationWindowDefinition
GroupAdd, anyPopupWindowDefinition                ,ahk_group fatalError1WindowDefinition
GroupAdd, anyPopupWindowDefinition                ,ahk_group fatalError2WindowDefinition
GroupAdd, anyPopupWindowDefinition                ,ahk_group projectwiseSavePromptWindowDefinition
GroupAdd, anyPopupWindowDefinition                ,ahk_group drawingFileNotValidWindowDefinition
GroupAdd, anyPopupWindowDefinition                ,ahk_group openDrawingErrorsFoundWindowDefinition



Loop{
    WinWait, ahk_group anyPopupWindowDefinition
    ; writeToLog("encountered a halting popup window that we will attempt to slap down.")
    ; might consider logging some of the window porperties or window text, etc.
    WinGet,idOfOffendingWindow,
    WinGet,processNameOfOffendingWindow,ProcessName,ahk_id %idOfOffendingWindow%
    writeToLog("idOfOffendingWindow: " . idOfOffendingWindow)
    writeToLog("processNameOfOffendingWindow: " . processNameOfOffendingWindow)

    if WinExist("ahk_group lostSheetSetAssociationWindowDefinition"){
        writeToLog("slapping down sheetset association error")
        ControlClick,Ignore the lost association,
    } else if WinExist("ahk_group fatalError1WindowDefinition") {
        writeToLog("slapping down fatal error popup 1")
        ControlSend, , {Enter}
    } else if WinExist("ahk_group fatalError2WindowDefinition") {
        writeToLog("slapping down fatal error popup 2")
        ControlSend, , !n
        WinWaitClose
        forcefullyKillAcad()

        ; I am including the force killing of autocad here because I suspect
        ; that it may be necessary to work around the powershell script's
        ; tendency to, when the powershell script goes to "wait-process" or
        ; doing Start-Process with the -Wait flag, which I think tries to wait
        ; for the process and all of its descendents to die, autocad's fatal
        ; crash leaves descendents still running, which causes a race condition
        ; with powershell waiting indefinitely because the descendents (like
        ; acwebbrowser, among others) have not yet died (And will never die
        ; unlkess killed)
        ;
        ; (due to a change I made in the way the powershell script waits (using
        ; "wait-process" instead of the -Wait flag to start-process), the forc
        ; killing of acad might no longer be necessary here (because I think that
        ; wait-process only waits for the process itself to die, not all of its
        ; descendents too, whereas the -Wait flag on Start-Process waits for the
        ; death of the process and all its descendents.  I could be confused
        ; about this whole interaction, but, empirically, it doesn't seem to
        ; hurt anything to perform a force-killing of acad here.

    } else if WinExist("ahk_group projectwiseSavePromptWindowDefinition") {
        writeToLog("slapping down projectwise save prompt popup")
        ControlSend, , !n
    } else if WinExist("ahk_group drawingFileNotValidWindowDefinition") {
        writeToLog("slapping down drawingFileNotValid popup")
        ControlSend, , {Enter}
        WinWaitClose
        forcefullyKillAcad()
    } else if WinExist("ahk_group openDrawingErrorsFoundWindowDefinition") {
        writeToLog("slapping down openDrawingErrorsFound popup")
        ; ControlSend, , !n
        ControlClick,No,

        ; there are two types of slapdown cases: one where our slapping down 
        ; will simply facilitate acad's inevitable death so we can cut our losses
        ; and mvoe on to the next file to be processed, and another where our slapping down 
        ; will goad autocad into continue with the processing. 
        ; this is noe of the latter. 
    
    } else { 
        writeToLog(""
            . "encountered an unexpected condition: "
            . "anyPopupWindowDefinition matches, "
            . "but none of the individual groups seem to match.  ")

        ; In this case, it is likely that the triggering window still exists, because we have not attempted to slap it down
        ; so continuing looping is probabaly going to immediately retriggerr, ad infinitum.
        ; therefore we wait for the triggering window to disappear (which will probably require manual action to happen, 
        ; but at least we wont be looping away clogging the log).
        WinWaitClose
    }

    ; in fact, its probably not a bad idea to ALWAYS wait for the triggering window 
    ; to close, because our whole goal is to slap down the window and make it close.  
    ; the only reason that we might not want to do winwaitclose is if it were in 
    ; some situations to perform the slapdown attempt multiple times to work. 
    ; but those cases should probably be handled within the individual 
    ; "else if" clauses above anyway.

    WinWaitClose

    Sleep, 500
    ; a bit of delay for good measure.
}

writeToLog(message){
    ; logFile:="$pathOfLogFile"
    logFile:="$($pathOfTemporarySlapdownLogDirectory)/" . A_ScriptName . ".log"
    timestampFormat:="yyyy/MM/dd HH:mm:ss"
    FormatTime, now,,%timestampFormat%
    FileAppend, %now%: %A_ScriptName% : %message%``n,%logFile%
}

forcefullyKillAcad(){
    writeToLog("forcefully killing autocad")
    Run, taskkill /t /f /im acad.exe, , Hide
    Run, taskkill /t /f /im acwebbrowser.exe, , Hide
}

"@
$pathOfPopupSlapdownScriptFile = join-path $pathOfMyOwnTempDirectory "popup_slapdown.ahk"
Set-Content -NoNewLine -Encoding ascii -Path $pathOfPopupSlapdownScriptFile -Value $popupSlapdownScriptContent


$businessScriptContent = @"
LOGFILEMODE 1
LOGFILEPATH $pathOfTemporaryAcadConsoleLogDirectory
audit Y
-purge regapps * N
-purge all * N
qsave 
QUIT 
"@

$pathOfBusinessScriptFile=join-path $pathOfMyOwnTempDirectory "business_script.scr"
Set-Content -NoNewLine -Encoding ascii -Path $pathOfBusinessScriptFile -Value $businessScriptContent

$setupScriptContent = @"
EXPERT 5
SSMAUTOOPEN 0
DEMANDLOAD 2
SAVETIME 0
FILEDIA 0
PROXYSHOW 0
PROXYNOTICE 0
reporterror 0
LOGFILEMODE 1
LOGFILEPATH $pathOfTemporaryAcadConsoleLogDirectory
STARTUP 0
QUIT Y 
"@

$pathOfSetupScriptFile=join-path $pathOfMyOwnTempDirectory "setup_script.scr"
Set-Content -NoNewLine -Encoding ascii -Path $pathOfSetupScriptFile -Value $setupScriptContent




$obscuringSuffix="220d09e398c24"
$dwgExtension=".dwg"

# $obscuredDwgExtension="$($nonobscuredDwgExtension)$obscuringSuffix"
# "obscuring" means changing the file extension of any file that we want to
# prevent AutoCAD from loading. Setting the doObscureNonactiveDwgFiles flag
# causes us to always (during the processing) keep all dwg files obscured except
# the dwg file that we are currently processing, with the idea that this
# prevents AutoCAD from trying to load any referenced files and potentially
# speeds up the processing and prevents some crashes. update: obscuring now
# means simply ensuring that the path ends with the obscuring suffix (which is
# assumed to be unique enough not to match any file extensions encountered in
# real life, although this is a bit of a lazy assumption, but good encough for
# the prupose at hand.
# we also assume that for each obscured, unobscured pair, only one of those files actually exists.


function getUnobscuredPath($pathOfFile){
    #takes a path that can be either obscured or unobscured, and returns the
    #unobscured version

    # $obscuredPath = $pathOfDwgFileToProcess + $obscuringSuffix 
    # I wish I had python's pathlib.Path class here.  There is probably something
    # similar in powershell or .net, but I don't know it off the top of
    # my head -- hence the ugly concatenation, above, to compute the
    # obscured path.
    # Actually, Split-Path covers most of the ground that I need.

    # $parent    = split-path -Path $pathOfFile -Parent
    # $extension = split-path -Path $pathOfFile -Extension
    # $leafBase  = split-path -Path $pathOfFile -LeafBase
    
    if ($pathOfFile.EndsWith($obscuringSuffix) ){
        $pathOfFile.Remove($pathOfFile.Length - $obscuringSuffix.Length)
    } else {
        $pathOfFile
    }
}


function getObscuredPath($pathOfFile){
    #takes a path that can be either obscured or unobscured, and returns the
    #obscured version
    (getUnobscuredPath $pathOfFile) + $obscuringSuffix
}

function obscure($pathOfFile){
    # it doesn't actually matter if the file is currently obscured or
    # unobscured. attempting the move (and ignoring any error thrown due to the
    # file already being in the desired obscurity state) suffices for our
    # purposes (although it will suppress some errors not caused by the file
    # already being in the desired state, but it is good enough for government
    # work.
    Move-Item (getUnobscuredPath $pathOfFile) (getObscuredPath $pathOfFile) -ErrorAction SilentlyContinue
}


function unobscure($pathOfFile){
    Move-Item (getObscuredPath $pathOfFile) (getUnobscuredPath $pathOfFile) -ErrorAction SilentlyContinue
}

function startAcadProcess($argumentList, $pathOfScriptFile=$null, $pathOfDwgFile=$null, $additionalArgs=$null, $strategy=$null){
    
    if($null -eq $strategy){
        # $strategy = "guiAcad"
        $strategy = "accoreconsole"
        # $strategy = "accoreconsole_with_embedded_open_command"
    }
    
    
    $argumentList = @()
    $s = @{}


    if($strategy -eq "guiAcad"){
        
        $argumentList += "/nologo"
        $argumentList += "/nossm"
        if($null -ne $pathOfScriptFile){
            $argumentList +=  @("/b", "`"$pathOfScriptFile`"")
        }
        if($null -ne $pathOfDwgFile){
            $argumentList +=  "`"$pathOfDwgFile`""
        }
        if($null -ne $additionalArgs){
            $argumentList +=  $additionalArgs
        }
        
        $s['FilePath']=$pathOfAcadExecutable 
        $s['WorkingDirectory']=$pathOfWorkingDirectoryInWhichToRunAcad
        $s['WindowStyle']="Minimized"

        # $s['WindowStyle']="Hidden"
        
    } elseif($strategy -eq "accoreconsole"){
        if($null -ne $pathOfScriptFile){
            $argumentList +=  @("/s", "`"$pathOfScriptFile`"")
        }
        if($null -ne $pathOfDwgFile){
            $argumentList +=  @( "/i", "`"$pathOfDwgFile`"")
        }
        if($null -ne $additionalArgs){
            $argumentList +=  $additionalArgs
        }
        $s['FilePath']=$pathOfAcadCoreConsoleExecutable 
        
        # $s['WorkingDirectory']=$pathOfWorkingDirectoryInWhichToRunAcad
        $s['WorkingDirectory']=$pathOfTemporaryAcadConsoleLogDirectory
        # Ithink accoreconsole insists on writing logs to the working directory. 
        # perhaps we should be capturing stdout and stderr instead of sweeping a log file.
        # $s['WindowStyle']="Minimized"
        # $s['WindowStyle']="Hidden"
        $s['NoNewWindow'] = 1
        
    } elseif($strategy -eq "accoreconsole_with_embedded_open_command"){
        $pathOfModifiedScriptFile = join-path $pathOfMyOwnTempDirectory "$(getPlainGuidString).scr"
        $modifiedScriptContent = ""

        if($null -ne $pathOfScriptFile){
            # $argumentList +=  @("/s", "`"$pathOfModifiedScriptFile`"")
            $s['RedirectStandardInput'] = $pathOfModifiedScriptFile
        }
        if($null -ne $pathOfDwgFile){
            $modifiedScriptContent = "OPEN $pathOfDwgFile" + "`r`n"
        }
        if($null -ne $additionalArgs){
            $argumentList +=  $additionalArgs
        }
        $modifiedScriptContent += (Get-Content -Raw -Path $pathOfScriptFile)
        Set-Content -NoNewLine -Encoding ascii -Path $pathOfModifiedScriptFile -Value $modifiedScriptContent

        $s['FilePath']=$pathOfAcadCoreConsoleExecutable 
        
        # $s['WorkingDirectory']=$pathOfWorkingDirectoryInWhichToRunAcad
        $s['WorkingDirectory']=$pathOfTemporaryAcadConsoleLogDirectory
        # Ithink accoreconsole insists on writing logs to the working directory. 
        # perhaps we should be capturing stdout and stderr instead of sweeping a log file.
        # $s['WindowStyle']="Minimized"
        # $s['WindowStyle']="Hidden"
        $s['NoNewWindow'] = 1
        
    } else {
        writeToLog "unrecognized strategy: $strategy"
    }
    
    # $s = @{
    #     FilePath=$pathOfAcadCoreConsoleExecutable 
    #     ArgumentList = $argumentList + @(
    #         "/nologo",
    #         "/nossm"
    #         # both nossm and nologo are an attempt to minimize acad from stealing focus.
    #     )
    #     WorkingDirectory=$pathOfWorkingDirectoryInWhichToRunAcad
    #     #  WindowStyle="Minimized"
    #     WindowStyle="Hidden"
    # }; 
    $s['ArgumentList'            ] = $argumentList 
    $s['RedirectStandardOutput'  ] = $pathOfStdoutCollectorFile
    $s['RedirectStandardError'   ] = $pathOfStderrCollectorFile
    Set-Content -NoNewLine -Encoding ascii -Path $pathOfStdoutCollectorFile -Value ""
    Set-Content -NoNewLine -Encoding ascii -Path $pathOfStderrCollectorFile -Value ""
    writeToLog ("command: `"$($s['FilePath'])`" $($argumentList -join " ")")
    $process = Start-Process -PassThru @s 
    $script:acadProcess = $process
    $process
}


function processSingleDwgFile($pathOfDwgFileToProcess, $pathsOfDwgFilesToObscure){
    #returns a dwgProcessingResult object


    $acadConsoleLogFileContent   = ""
    $acadErrFileContent          = ""
    $stdErrContent               = ""
    $stdOutContent               = ""
    $slapdownLogFileContent      = ""

     


    writeToLog  "Now working on dwg file $pathOfDwgFileToProcess"
    $processingDuration          = (New-TimeSpan )        

    if ($doObscureNonactiveDwgFiles){
        foreach ($pathOfDwgFile in $pathsOfDwgFilesToObscure){obscure($pathOfDwgFile)}
        unobscure($pathOfDwgFileToProcess)
    }

    # clear the read-only flag, if it is present, to avoid AutoCAD prompting about opening read only
    $isReadOnly = (Get-ItemProperty -Path $pathOfDwgFileToProcess -Name IsReadOnly).IsReadOnly
    # writeToLog ("isReadOnly: $isReadOnly")
    If ($isReadOnly){
        writeToLog ("clearing the readOnly flag on $pathOfDwgFileToProcess")
        Set-ItemProperty -Path $pathOfDwgFileToProcess -Name IsReadOnly -Value $false
    }
            
    $initialFileSize = (Get-Item $pathOfDwgFileToProcess).length
    $initialFileHash = (Get-FileHash -Algorithm SHA1 $pathOfDwgFileToProcess).Hash
    
    $startTime = Get-Date        
    $process = startAcadProcess -pathOfScriptFile $pathOfBusinessScriptFile -pathOfDwgFile $pathOfDwgFileToProcess
    $process | Wait-Process
    $endTime = Get-Date
    $processingDuration += $endTime - $startTime

    # give some time for autohotkey slapdowns to finish forcefully killing acad (useful in case of fatal error)
    Start-Sleep -Seconds 1

    $acadConsoleLogFileContent += (flushAcadConsoleLog)
    $stdErrContent             += (flushStdErrCollectorFile)
    $stdOutContent             += (flushStdOutCollectorFile)
    $slapdownLogFileContent    += (flushSlapdownLog)
    $acadErrFileContent        += (flushAcadErrFile)

    
    $finalFileSize = (Get-Item $pathOfDwgFileToProcess).length
    $finalFileHash = (Get-FileHash -Algorithm SHA1 $pathOfDwgFileToProcess).Hash

    if ($initialFileHash -eq $finalFileHash){
        writeToLog "The first attempt at processing the file produced no change.  Therefore, we will try again in a different way."

        $startTime = Get-Date   
        $process = startAcadProcess -pathOfScriptFile $pathOfBusinessScriptFile -pathOfDwgFile $pathOfDwgFileToProcess -strategy "guiAcad"
        $process | Wait-Process
        $endTime = Get-Date
        $processingDuration += $endTime - $startTime

        # wait for autohotkey slapdowns to finish forcefully killing acad (useful in case of fatal error)
        Start-Sleep -Seconds 3

        $acadConsoleLogFileContent += "`n" + (flushAcadConsoleLog)
        $stdErrContent             += "`n" + (flushStdErrCollectorFile)
        $stdOutContent             += "`n" + (flushStdOutCollectorFile)
        $slapdownLogFileContent    += "`n" + (flushSlapdownLog)
        $acadErrFileContent        += "`n" + (flushAcadErrFile)
        
        $finalFileSize = (Get-Item $pathOfDwgFileToProcess).length
        $finalFileHash = (Get-FileHash -Algorithm SHA1 $pathOfDwgFileToProcess).Hash
    }


    if ($doObscureNonactiveDwgFiles){
        foreach ($pathOfDwgFile in $pathsOfDwgFilesToObscure){unobscure($pathOfDwgFile)}
        unobscure($pathOfDwgFileToProcess)
    }

    # writeToLog "finished processing file $($i + 1) in $(toHumanReadableDuration($endTime - $startTime)): $pathOfDwgFileToProcess"
    
    # writeToLog ("file $($i + 1) size " + "$(if($finalFileSize -eq $initialFileSize){ "remained constant at $(toHumanReadableDataSize $initialFileSize) ."  } else {"changed from $(toHumanReadableDataSize $initialFileSize) to $(toHumanReadableDataSize $finalFileSize) ."})")
    # if ($initialFileHash -eq $finalFileHash){
    #     writeToLog "file $($i +1) did not change.  hash remained constant at $initialFileHash"
    # } else {
    #     writeToLog "file $($i +1) hash changed from $initialFileHash to $finalFileHash"
    # }

    $returnValue = [PSCustomObject]@{
        pathOfDwgFile              = $pathOfDwgFileToProcess
        acadConsoleLogFileContent  = $acadConsoleLogFileContent
        acadErrFileContent         = $acadErrFileContent
        stdErrContent              = $stdErrContent
        stdOutContent              = $stdOutContent
        slapdownLogFileContent     = $slapdownLogFileContent
        processingDuration         = $processingDuration
        initialFileSize            = $initialFileSize
        finalFileSize              = $finalFileSize
        initialFileHash            = $initialFileHash
        finalFileHash              = $finalFileHash
    }      
    $returnValue
}

# function toChangeClause($initialValue, $finalValue, $formatter={param([String]$arg0); $arg0}){
function toChangeClause($initialValue, $finalValue, $formatter="write-output"){
    # if($initialValue -eq $finalValue){
    #     "remained constant at $($formatter.Invoke($initialValue))"
    # } else {
    #     "changed from $($formatter.Invoke($initialValue)) to $($formatter.Invoke($finalValue))"
    # }
    if($initialValue -eq $finalValue){
        "remained constant at $(& $formatter $initialValue)"
    } else {
        "changed from $(& $formatter $initialValue) to $(& $formatter $finalValue)"
    }
}


function processingResultToReport($result, $i){
    # $startTime = Get-Date 
    $m = ""

    $m += "file $($i + 1) processingDuration:               "
    $m += "$(toHumanReadableDuration $result.processingDuration)" + "`n"

    $m += "file $($i + 1) size:                             "
    $m += (toChangeClause $result.initialFileSize $result.finalFileSize toHumanReadableDataSize) + "`n"

    $m += "file $($i + 1) hash:                             "
    $m += (toChangeClause $result.initialFileHash $result.finalFileHash) + "`n"

    $pathOfReviewableAcadConsoleLogFile = ((join-path $pathOfReviewableAcadConsoleLogDirectory ([IO.Path]::GetRelativePath($pathOfDirectoryInWhichToSearchForDwgFiles, $pathOfDwgFileToProcess))) + "." + (getPlainGuidString) + ".console_log")
    $pathOfReviewableAcadStdoutLogFile = ((join-path $pathOfReviewableAcadConsoleLogDirectory ([IO.Path]::GetRelativePath($pathOfDirectoryInWhichToSearchForDwgFiles, $pathOfDwgFileToProcess))) + "." + (getPlainGuidString) + ".stdout_log")
    
    New-Item -ItemType "directory" -Path (Split-Path -Path $pathOfReviewableAcadConsoleLogFile -Parent)  -ErrorAction SilentlyContinue | Out-Null    
    New-Item -ItemType "directory" -Path (Split-Path -Path $pathOfReviewableAcadStdoutLogFile -Parent) -ErrorAction SilentlyContinue | Out-Null    

    Add-Content -Path $pathOfReviewableAcadConsoleLogFile -Value $result.acadConsoleLogFileContent
    Add-Content -Path $pathOfReviewableAcadStdoutLogFile -Value $result.stdOutContent

    # $m +=  "file $($i + 1) acadConsoleLogFileContent:`n$(indent $result.acadConsoleLogFileContent -doLineNumbers $true)" + "`n"
    $m += "file $($i + 1) acadConsoleLogFile:               $pathOfReviewableAcadConsoleLogFile" + "`n"
    $m += "file $($i + 1) acadConsoleLogFileContent.Length: $($result.acadConsoleLogFileContent.Length)" + "`n"

    # $m +=  "file $($i + 1) stdOutContent:`n$(indent $result.stdOutContent -doLineNumbers $true)" + "`n"
    $m += "file $($i + 1) stdOutLogFile:                    $pathOfReviewableAcadStdoutLogFile" + "`n"
    $m += "file $($i + 1) stdOutContent.Length:             $($result.stdOutContent.Length)" + "`n"


    $m += "file $($i + 1) acadErrFileContent:               $( if($result.acadErrFileContent.Length -eq 0 ){ "(empty)" } else { "`n(indent $($result.acadErrFileContent) -doLineNumbers $true)" } )" + "`n"
    # $m += "file $($i + 1) stdErrContent:`n$(indent $result.stdErrContent -doLineNumbers $true)" + "`n"
    $m += "file $($i + 1) stdErrContent:                    $( if($result.stdErrContent.Length -eq 0 ){ "(empty)" } else { "`n(indent $($result.stdErrContent) -doLineNumbers $true)" } )" + "`n"
    # $m += "file $($i + 1) slapdownLogFileContent:`n$(indent $result.slapdownLogFileContent -doLineNumbers $true)" + "`n"
    $m += "file $($i + 1) slapdownLogFileContent:           $( if($result.slapdownLogFileContent.Length -eq 0 ){ "(empty)" } else { "`n(indent $($result.slapdownLogFileContent) -doLineNumbers $true)" } )"
    # $endTime = Get-Date 
    # writeToLog "generated report in $(toHumanReadableDuration($endTime - $startTime))"
    $m
}

writeToLog "pathOfLogFile:                                      $pathOfLogFile"
writeToLog "pathOfAcadExecutable:                               $pathOfAcadExecutable"
writeToLog "pathOfAcadCoreConsoleExecutable:                    $pathOfAcadCoreConsoleExecutable"
writeToLog "pathOfDirectoryInWhichToSearchForDwgFiles:          $pathOfDirectoryInWhichToSearchForDwgFiles"
writeToLog "pathOfTemporaryAcadConsoleLogDirectory:             $pathOfTemporaryAcadConsoleLogDirectory"
writeToLog "pathOfTemporarySlapdownLogDirectory:                $pathOfTemporarySlapdownLogDirectory"
writeToLog "pathOfBusinessScriptFile:                           $pathOfBusinessScriptFile"
writeToLog "pathOfSetupScriptFile:                              $pathOfSetupScriptFile"
writeToLog "pathOfPopupSlapdownScriptFile:                      $pathOfPopupSlapdownScriptFile"
writeToLog "pathOfWorkingDirectoryInWhichToRunAcad:             $pathOfWorkingDirectoryInWhichToRunAcad"
writeToLog "PSScriptRoot:                                       $PSScriptRoot"
writeToLog "PSVersionTable.PSVersion:                           $($PSVersionTable.PSVersion)"
writeToLog "skipToOrdinal:                                      $skipToOrdinal"     
writeToLog "doObscureNonactiveDwgFiles:                         $doObscureNonactiveDwgFiles"     
writeToLog "pathOfStdoutCollectorFile:                          $pathOfStdoutCollectorFile"     
writeToLog "pathOfStderrCollectorFile:                          $pathOfStderrCollectorFile"     
writeToLog "pathOfReviewableAcadConsoleLogDirectory:            $pathOfReviewableAcadConsoleLogDirectory"     
writeToLog "setupScriptContent:`n$(indent $setupScriptContent -doLineNumbers $true)"     
writeToLog "businessScriptContent:`n$(indent $businessScriptContent -doLineNumbers $true)"     


# obscuringSuffix is a a a suffix that we can stick on the file extension to effectively hide the file from autocad, 
# in order to prevent autocad from loading the file as an xref.
# $pathsOfDwgFilesToProcess = @(Get-ChildItem -Path $pathOfDirectoryInWhichToSearchForDwgFiles -File -Recurse | Where-Object {$_.Extension -eq ".dwg"} | foreach-object {$_.FullName})
# $pathsOfDwgFilesToProcess = @( 
#     Get-ChildItem -Path $pathOfDirectoryInWhichToSearchForDwgFiles -File -Recurse `
#         | Where-Object {$_.Extension -in @($nonobscuredDwgExtension, $obscuredDwgExtension)} `
#         | foreach-object {$_.FullName}
# )
# we are processing "obscured" in case they are left over from an aborted previous run of theis script.
## actually not bothering with that.
# $pathsOfDwgFilesToProcess = @( 
#     Get-ChildItem -Path $pathOfDirectoryInWhichToSearchForDwgFiles -File -Recurse `
#         | Where-Object {$_.Extension -eq $nonobscuredDwgExtension} `
#         | foreach-object {$_.FullName}
# )

$pathsOfDwgFilesToProcess = @( 
    Get-ChildItem -Path $pathOfDirectoryInWhichToSearchForDwgFiles -File -Recurse `
        | foreach-object {getUnobscuredPath $_.FullName} `
        | Where-Object {(split-path $_ -Extension) -eq $dwgExtension}
)

# $pathsOfDwgFilesToProcess will always contain the unobscured form of the paths.



$message = "We will now process $($pathsOfDwgFilesToProcess.length) dwg files: " + "`n"
$message += indent ( 
    @( 
        ( 1 .. $pathsOfDwgFilesToProcess.length ) 
        | foreach-object { "file $($_): $($pathsOfDwgFilesToProcess[$_ - 1])" } 
    ) -join "`n" 
)
writeToLog  $message


Try {
    if($pathsOfDwgFilesToProcess){
        $overallStartTime = Get-Date
        $countOfProcessedFiles = 0
        $totalInitialFileSize = 0
        $totalFinalFileSize = 0
        $totalFirstPassProcessingDuration = New-TimeSpan
        $totalSecondPassProcessingDuration = New-TimeSpan


        if ($doObscureNonactiveDwgFiles){
            writeToLog  "Now ensuring that all dwg files are obscured."
            foreach ($pathOfDwgFileToProcess in $pathsOfDwgFilesToProcess) {obscure($pathOfDwgFileToProcess)}
        }

        
        
        sweepAcadErrFileToOurLog 
        
        # it would be better to just delete any existing
        # error file rather than sweeping it to our log, because the error file, if
        # it exists, would not be related to what we are doing in this invokaction
        # of the script, because we have not yet run autocad (in this invokation).
        # the goal is to get rid of any existing error file toa void inadvertently
        # sweeping an old, irrelevant, error file into our log with the calls to
        # sweepAcadErrFileToOurLog, below.
        # but betterr to sweep here then to do nothing at all to delete any existing error file.


        writeToLog  "Now starting the popup slapdown script"
        $script:popupSlapdownScriptProcess = Start-Process -FilePath $pathOfPopupSlapdownScriptFile -PassThru
        
        writeToLog  "Now setting AutoCAD system variables via registry"
        setRegistryValueInAllAcadProfiles -relativePathAndName "Variables\STARTUP" -value "0"
        setRegistryValueInAllAcadProfiles -relativePathAndName "Variables\STARTMODE" -value "0"

        writeToLog  "Now running the autoCAD setup script."
        $process = startAcadProcess -pathOfScriptFile $pathOfSetupScriptFile
        
        $process | Wait-Process
        sweepAcadConsoleLogToOurLog   
        sweepAcadErrFileToOurLog
        sweepSlapdownLogToOurLog
        sweepStandardStreamCollectorFilesToOurLog



        for ($i=0; $i -lt $pathsOfDwgFilesToProcess.length; $i++){
            if ( ($skipToOrdinal -ne $null)  -and ( ($i + 1) -lt  $skipToOrdinal)  ){
                writeToLog "skipping until we reach ordinal $skipToOrdinal"
            } else {          
                $pathOfDwgFileToProcess = $pathsOfDwgFilesToProcess[$i]
                    
                writeToLog  "Now working on dwg file $($i + 1) of $($pathsOfDwgFilesToProcess.length): $pathOfDwgFileToProcess"

                $result1 = (processSingleDwgFile -pathOfDwgFileToProcess $pathOfDwgFileToProcess -pathsOfDwgFilesToObscure $pathsOfDwgFilesToProcess)
                writeToLog "file $($i + 1) first pass processing report:  `n$(indent (processingResultToReport $result1 $i))"


                $result2 = (processSingleDwgFile -pathOfDwgFileToProcess $pathOfDwgFileToProcess -pathsOfDwgFilesToObscure $pathsOfDwgFilesToProcess)
                writeToLog "file $($i + 1) second pass processing report:  `n$(indent (processingResultToReport $result2 $i))"

                
                $totalInitialFileSize += $result1.initialFileSize
                $totalFinalFileSize += $result2.finalFileSize
                $totalFirstPassProcessingDuration += $result1.processingDuration
                $totalSecondPassProcessingDuration += $result2.processingDuration


                $m = "" 
                $m += "file $($i + 1) processingDuration "
                $m += (toChangeClause $result1.processingDuration $result2.processingDuration toHumanReadableDuration) + "."
                writeToLog $m

                $m = "" 
                $m += "file $($i + 1) size "
                $m += (toChangeClause $result1.initialFileSize $result2.finalFileSize toHumanReadableDataSize) + "."
                writeToLog $m

                $countOfProcessedFiles = $countOfProcessedFiles + 1
            }
        }

    }

    writeToLog  "finished"
} Finally {
    #cleanup:
    Write-Host "cleaning up..."

    $script:popupSlapdownScriptProcess.Kill($True)
    $script:acadProcess.Kill($True)

    if ($doObscureNonactiveDwgFiles){
        writeToLog  "Now ensuring that all dwg files are non-obscured."
        foreach ($pathOfDwgFileToProcess in $pathsOfDwgFilesToProcess) {unobscure($pathOfDwgFileToProcess)}
    }

    $overallEndTime = Get-Date
    writeToLog "processed $countOfProcessedFiles files in $(toHumanReadableDuration($overallEndTime - $overallStartTime))."
    writeToLog "total size of processed files $(toChangeClause $totalInitialFileSize $totalFinalFileSize toHumanReadableDataSize)."
    writeToLog "between first and second passes, total processing duration $(toChangeClause $totalFirstPassProcessingDuration $totalSecondPassProcessingDuration toHumanReadableDuration)."
}


# == various AutoCAD prompts for user-intervention that tend to appear while
# running the script and how to deal with them: -- Some of these can be dealt
# with programmatically by fiddling with the Windows registry or running AutoCAD
# commands in the setup script, but others are not so easy to programatically
# prevent, and in those cases, at least one manual click is required to prevent
# recurrence.
# - missing SHX file -tick the "don't show me this dialog again" checkbox.
# - reset annotative scales? (paraphrase)
# - you should perform a recover operation before opening this file (paraphrase)
# - display proxy graphics ?
# - missing xrefs (paraphrase) (from my memory)
# - "Saving the drawing will update any AEC objects in it to the current version
#   which will be incompatible with earlier versions."
# - "The sheet set association has been lost. What do you want to do?"

