#!pwsh



$pathOfAcadExecutable = join-path $env:ProgramFiles "Autodesk/AutoCAD 2018/acad.exe"
$pathOfDirectoryInWhichToSearchForDwgFiles='C:\bms'
$pathOfWorkingDirectoryInWhichToRunAcad=$PSScriptRoot

$skipToOrdinal = $null
# a hack for resuming our place in the sequence, intended for use in debugging.
# unassign or set to $null to to not skip anything


# override some of the default parameters, on Neil's machine only:
if($env:computername -match "AUTOSCAN-WK07W7"){
    $pathOfAcadExecutable = join-path $env:ProgramFiles "Autodesk\AutoCAD 2023\acad.exe"
    # $pathOfDirectoryInWhichToSearchForDwgFiles='J:\Nakano\2022-02-08_projectwise_aec_acad_troubleshoot\testroot'
    $pathOfDirectoryInWhichToSearchForDwgFiles='C:\work\nakano_marginal_way'
} 



# $pathOfLogFile = (join-path $env:TEMP $MyInvocation.MyCommand.Name) + ".log"
$pathOfLogFile = (join-path $pathOfDirectoryInWhichToSearchForDwgFiles $MyInvocation.MyCommand.Name) + ".log"
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




writeToLog "$($MyInvocation.MyCommand.Name) is running."
writeToLog "pathOfLogFile: $pathOfLogFile"



$pathOfTemporaryAcadConsoleLogDirectory = join-path $env:TEMP ((New-Guid).Guid)

function sweepAcadConsoleLogToOurLog(){
    writeToLog "sweeping acad console logs"
    foreach($acadConsoleLogFile in (Get-ChildItem -Path $pathOfTemporaryAcadConsoleLogDirectory -File -Recurse )){
        writeToLog "content of $($acadConsoleLogFile.FullName):`n$(Get-Content -Raw -Path $acadConsoleLogFile.FullName)"
        Remove-Item -Path $acadConsoleLogFile.FullName 
    }
}

function sweepAcadErrFileToOurLog(){
    writeToLog "sweeping acad err file"
    $pathOfAcadErrFile=Join-Path $pathOfWorkingDirectoryInWhichToRunAcad "acad.err"
    if(Test-Path -Path $pathOfAcadErrFile -PathType Leaf){
        writeToLog "content of $($pathOfAcadErrFile):`n$(Get-Content -Raw -Path $pathOfAcadErrFile)"
        Remove-Item -Path $pathOfAcadErrFile
    } else {
        writeToLog "hooray!  $pathOfAcadErrFile does not exist."
    }
}


New-Item -ItemType "directory" -Path $pathOfTemporaryAcadConsoleLogDirectory | Out-Null

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
        ControlSend, , !n
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
    logFile:="$pathOfLogFile"
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
$pathOfPopupSlapdownScriptFile=join-path $env:TEMP ("popup_slapdown" + ".ahk")
Set-Content -NoNewLine -Encoding ascii -Path $pathOfPopupSlapdownScriptFile -Value $popupSlapdownScriptContent


$businessScriptContent = @"
audit Y
-purge regapps * N
-purge all * N
qsave 
QUIT 
"@

$pathOfBusinessScriptFile=join-path $env:TEMP ((New-Guid).Guid + ".scr")
Set-Content -NoNewLine -Encoding ascii -Path $pathOfBusinessScriptFile -Value $businessScriptContent

$setupScriptContent = @"
EXPERT 5
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

$pathOfSetupScriptFile=join-path $env:TEMP ((New-Guid).Guid + ".scr")
Set-Content -NoNewLine -Encoding ascii -Path $pathOfSetupScriptFile -Value $setupScriptContent



$doObscureNonactiveDwgFiles = $True
$obscuringSuffix="220d09e398c24"
$dwgExtension=".dwg"
$nonobscuredDwgExtension=$dwgExtension
$obscuredDwgExtension="$($nonobscuredDwgExtension)$obscuringSuffix"
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





writeToLog  "pathOfAcadExecutable:                               $pathOfAcadExecutable"
writeToLog  "pathOfDirectoryInWhichToSearchForDwgFiles:          $pathOfDirectoryInWhichToSearchForDwgFiles"
writeToLog  "pathOfTemporaryAcadConsoleLogDirectory:             $pathOfTemporaryAcadConsoleLogDirectory"
writeToLog  "pathOfBusinessScriptFile:                           $pathOfBusinessScriptFile"
writeToLog  "pathOfSetupScriptFile:                              $pathOfSetupScriptFile"
writeToLog  "pathOfPopupSlapdownScriptFile:                      $pathOfPopupSlapdownScriptFile"
writeToLog  "pathOfWorkingDirectoryInWhichToRunAcad:             $pathOfWorkingDirectoryInWhichToRunAcad"
writeToLog  "PSScriptRoot:                                       $PSScriptRoot"
writeToLog  "PSVersionTable.PSVersion:                           $($PSVersionTable.PSVersion)"
writeToLog  "skipToOrdinal:                                      $skipToOrdinal"     
writeToLog  "doObscureNonactiveDwgFiles:                         $doObscureNonactiveDwgFiles"     


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


# writeToLog  "We will now process $($pathsOfDwgFilesToProcess.length) dwg files."
$message = "We will now process $($pathsOfDwgFilesToProcess.length) dwg files: "
for ($i=0; $i -lt $pathsOfDwgFilesToProcess.length; $i++){
    $message = $message + "`n"
    $message = $message + "    file " + ($i + 1) + ": " + $pathsOfDwgFilesToProcess[$i]
}
writeToLog  $message



if($pathsOfDwgFilesToProcess){
    $overallStartTime = Get-Date
    $countOfProcessedFiles = 0

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
    Start-Process -FilePath $pathOfPopupSlapdownScriptFile
    
    writeToLog  "Now setting AutoCAD system variables via registry"
    setRegistryValueInAllAcadProfiles -relativePathAndName "Variables\STARTUP" -value "0"
    setRegistryValueInAllAcadProfiles -relativePathAndName "Variables\STARTMODE" -value "0"

    writeToLog  "Now running the autoCAD setup script."
    # Start-Process -Wait -FilePath $pathOfAcadExecutable -ArgumentList @("/b", "`"$pathOfSetupScriptFile`"")
    
    $s = @{
        FilePath=$pathOfAcadExecutable 
        ArgumentList=@(
            "/b", "`"$pathOfSetupScriptFile`""
            "/nologo"
        )
        WorkingDirectory=$pathOfWorkingDirectoryInWhichToRunAcad
        WindowStyle="Minimized"
    }; 
    $process = Start-Process -PassThru @s 
    # Write-Host "process: ";    $process | fl
    $process | Wait-Process
    sweepAcadConsoleLogToOurLog   
    sweepAcadErrFileToOurLog





    for ($i=0; $i -lt $pathsOfDwgFilesToProcess.length; $i++){
        if ( ($skipToOrdinal -ne $null)  -and ( ($i + 1) -lt  $skipToOrdinal)  ){
            writeToLog "skipping until we reach ordinal $skipToOrdinal"
        } else {
        
            $startTime = Get-Date
            
            $pathOfDwgFileToProcess = $pathsOfDwgFilesToProcess[$i]
            writeToLog  "Now working on dwg file $($i + 1) of $($pathsOfDwgFilesToProcess.length): $pathOfDwgFileToProcess"
            
            
            if ($doObscureNonactiveDwgFiles){unobscure($pathOfDwgFileToProcess)}
            
            # & cmd /c "$pathOfAcadExecutable" "$pathOfDwgFileToProcess" /b "$pathOfScriptFile"
            
            # clear the read-only flag, if it is present, to avoid AutoCAD prompting about opening read only

            $isReadOnly = (Get-ItemProperty -Path $pathOfDwgFileToProcess -Name IsReadOnly).IsReadOnly
            # writeToLog ("isReadOnly: $isReadOnly")
            If ($isReadOnly){
                writeToLog ("clearing the readOnly flag on $pathOfDwgFileToProcess")
                Set-ItemProperty -Path $pathOfDwgFileToProcess -Name IsReadOnly -Value $false
            }



            # Start-Process -Wait -FilePath $pathOfAcadExecutable -ArgumentList @("`"$pathOfDwgFileToProcess`"", "/b", "`"$pathOfBusinessScriptFile`"")
            # $s = @{
            #     Wait=1
            #     FilePath=$pathOfAcadExecutable 
            #     ArgumentList=@("`"$pathOfDwgFileToProcess`"", "/b", "`"$pathOfBusinessScriptFile`"")
            #     WorkingDirectory=$pathOfWorkingDirectoryInWhichToRunAcad
            # }; 
            # $process = Start-Process @s
            

            
            $s = @{
                FilePath=$pathOfAcadExecutable 
                ArgumentList=@(
                    "`"$pathOfDwgFileToProcess`"", 
                    "/b", "`"$pathOfBusinessScriptFile`"",
                    "/nologo"
                )
                WorkingDirectory=$pathOfWorkingDirectoryInWhichToRunAcad
                WindowStyle="Minimized"
            }; 
            $process = Start-Process -PassThru @s 
            $process | Wait-Process


            # wait for autohotkey slapdowns to finish forcefully killing acad (useful in case of fatal error)
            Start-Sleep -Seconds 3

            sweepAcadConsoleLogToOurLog   
            sweepAcadErrFileToOurLog
            if ($doObscureNonactiveDwgFiles){obscure($pathOfDwgFileToProcess)}
            

            
            $endTime = Get-Date
            writeToLog "finished processing in $([math]::Round($($endTime - $startTime).TotalSeconds)) seconds: $pathOfDwgFileToProcess"

            $countOfProcessedFiles = $countOfProcessedFiles + 1
        }
    }



    if ($doObscureNonactiveDwgFiles){
        writeToLog  "Now ensuring that all dwg files are non-obscured."
        foreach ($pathOfDwgFileToProcess in $pathsOfDwgFilesToProcess) {unobscure($pathOfDwgFileToProcess)}
    }

    $overallEndTime = Get-Date

    writeToLog "processed $countOfProcessedFiles files in $([math]::Round($($overallEndTime - $overallStartTime).TotalSeconds)) seconds."

}

writeToLog  "finished"



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

