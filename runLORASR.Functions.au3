#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Functions
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.12
 Modified:       2017.08.30
 Version:        0.4.2.2

 Script Function:
	Functions used by runLORASR

#ce ----------------------------------------------------------------------------

#include-once
#include <Date.au3>
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>

; Global declarations
Global $g_iConsoleVerbosity = 5
Global $g_iLogFileVerbosity = 4
Global $g_iMessageVerbosity = 1
Global $g_sLogFile = "runLORASR.log"

; Create new log file on first open
CreateLogFile($g_sLogFile, @WorkingDir)
LogMessage("Loaded runLORASR.Functions version 0.4.2.2", 3)

; Function to read settings from runLORASR.ini file
Func GetSettings($sWorkingDirectory, ByRef $sProgramPath, ByRef $sSimulationProgram, ByRef $sSweepFile, ByRef $sTemplateFile, ByRef $sResultsFile, ByRef $sPlotFile, ByRef $sInputFolder, ByRef $sOutputFolder, ByRef $sRunFolder, ByRef $sIncompleteFolder, ByRef $bCleanup)
	LogMessage("Called GetSettings($sWorkingDirectory = " & $sWorkingDirectory & ", ... )", 5)

	; Declarations
	Local $sSettingsFile = "runLORASR.ini"
	Local $asSettingsSearchPath = ["C:\Program Files (x86)\LORASR", "C:\Program Files\LORASR", "C:\LORASR"]
	Local $sFoundFile = ""

	; Find INI file
	LogMessage("Searching for settings file " & $sSettingsFile, 3, "GetSettings")
	For $sSearchFolder in $asSettingsSearchPath
		$sFoundFile = FindFile($sSettingsFile, $sWorkingDirectory, $sSearchFolder, False)
		If $sFoundFile Then ExitLoop
	Next ; $sSearchFolder
	If $sFoundFile Then
		; Found settings file
		LogMessage("Using settings file at " & $sFoundFile, 4, "GetSettings")
		$sSettingsFile = $sFoundFile
	Else
		; File not found
		ThrowError("Could not find settings file " & $sSettingsFile & ". Switching to default values.", 3, "GetSettings", @error)
		SetError(1)
	EndIf

	; Load settings from INI file
	; On error the variables are set to the given default values
	$sProgramPath = IniRead($sSettingsFile, "Files and folders", "ProgramPath", "C:\Program Files (x86)\LORASR")
	LogMessage("ProgramPath = " & $sProgramPath, 4, "GetSettings")
	$sSimulationProgram = IniRead($sSettingsFile, "Files and folders", "SimulationProgram", "LORASR.exe")
	LogMessage("SimulationProgram = " & $sSimulationProgram, 4, "GetSettings")
	$sSweepFile = IniRead($sSettingsFile, "Files and folders", "SweepFile", "Sweep.xlsx")
	LogMessage("SweepFile = " & $sSweepFile, 4, "GetSettings")
	$sTemplateFile = IniRead($sSettingsFile, "Files and folders", "TemplateFile", "Template.txt")
	LogMessage("TemplateFile = " & $sTemplateFile, 4, "GetSettings")
	$sResultsFile = IniRead($sSettingsFile, "Files and folders", "ResultsFile", "Batch results.csv")
	LogMessage("ResultsFile = " & $sResultsFile, 4, "GetSettings")
	$sPlotFile = IniRead($sSettingsFile, "Files and folders", "PlotFile", "Plots.xlsx")
	LogMessage("PlotFile = " & $sPlotFile, 4, "GetSettings")
	$g_sLogFile = IniRead($sSettingsFile, "Files and folders", "LogFile", "runLORASR.log")
	LogMessage("LogFile = " & $g_sLogFile, 4, "GetSettings")
	$sInputFolder = IniRead($sSettingsFile, "Files and folders", "InputFolder", "Input")
	LogMessage("InputFolder = " & $sInputFolder, 4, "GetSettings")
	$sOutputFolder = IniRead($sSettingsFile, "Files and folders", "OutputFolder", "Output")
	LogMessage("OutputFolder = " & $sOutputFolder, 4, "GetSettings")
	$sRunFolder = IniRead($sSettingsFile, "Files and folders", "RunFolder", "Runs")
	LogMessage("RunFolder = " & $sRunFolder, 4, "GetSettings")
	$sIncompleteFolder = IniRead($sSettingsFile, "Files and folders", "IncompleteFolder", "Incomplete")
	LogMessage("IncompleteFolder = " & $sIncompleteFolder, 4, "GetSettings")

	$bCleanup = (StringCompare(IniRead($sSettingsFile, "Options", "Cleanup", "True"), "True") = 0)
	LogMessage("Cleanup = " & $bCleanup, 4, "GetSettings")
	$g_iConsoleVerbosity = Number(IniRead($sSettingsFile, "Options", "ConsoleVerbosity", "5"))
	LogMessage("ConsoleVerbosity = " & $g_iConsoleVerbosity, 4, "GetSettings")
	$g_iLogFileVerbosity = Number(IniRead($sSettingsFile, "Options", "LogFileVerbosity", "4"))
	LogMessage("LogFileVerbosity = " & $g_iLogFileVerbosity, 4, "GetSettings")
	$g_iMessageVerbosity = Number(IniRead($sSettingsFile, "Options", "MessageVerbosity", "1"))
	LogMessage("MessageVerbosity = " & $g_iMessageVerbosity, 4, "GetSettings")

	; Exit function
	LogMessage("All settings now set.", 3, "GetSettings")
	Return (Not @error)

EndFunc

; Function to find required files
Func FindFile($sFindFileName, $sWorkingDir = @WorkingDir, $sMasterDir = $sWorkingDir & "\Input", $bCopy = True)
	LogMessage("Called FindFile($sFindFileName = " & $sFindFileName & ", $sWorkingDir = " & $sWorkingDir & ", $sMasterDir = " & $sMasterDir & ", $bCopy = " & $bCopy & ")", 5)

	; Declarations
	Local $sFoundFile = ""

	; Look in working directory first
	If FileExists($sWorkingDir & "\" & $sFindFileName) Then
		LogMessage("File found in working directory", 5, "FindFile")
		$sFoundFile = $sWorkingDir & "\" & $sFindFileName
	Else
		; If not found, check master directory
		If FileExists($sMasterDir & "\" & $sFindFileName) Then
			LogMessage("File found in master directory", 5, "FindFile")
			If $bCopy Then
				LogMessage("Copying file to working directory", 4, "FindFile")
				; Copy the master to the working directory
				If FileCopy($sMasterDir & "\" & $sFindFileName, $sWorkingDir & "\" & $sFindFileName) Then
					; Report the copied file location
					$sFoundFile = $sWorkingDir & "\" & $sFindFileName
				Else
					; Return failure: copy failed
					ThrowError("Could not copy file from " & $sMasterDir & "\" & $sFindFileName & "to" & $sWorkingDir & "\" & $sFindFileName, 3, "FindFile", @error)
					SetError(2)
					Return 0
				EndIf
			Else
				; Report the master file location
				$sFoundFile = $sMasterDir & "\" & $sFindFileName
			EndIf
		Else
			; Return failure: file not found
			LogMessage("File " & $sFindFileName & " not found in working directory nor in " & $sMasterDir, 3, "FindFile")
			SetError(1)
			Return 0
		EndIf
	EndIf

	; Return result
	LogMessage("Found file = " & $sFoundFile, 3, "FindFile")
	Return $sFoundFile

EndFunc

; Function to copy a file or set of files
Func CopyFiles($sCopyFileName, $sCopySourceFolder, $sCopyDestinationFolder, $bOverwrite = False)
	LogMessage("Called CopyFiles($sCopyFileName = " & $sCopyFileName & ", $sCopySourceFolder = " & $sCopySourceFolder & ", $sCopyDestinationFolder = " & $sCopyDestinationFolder & ", $bOverwrite = " & $bOverwrite & ")", 5)

	Local $asCopyFiles
	Local $iCurrentFile
	Local $sCurrentFile
	Local $FLAGS

	; Should overwrite?
	If $bOverwrite Then
		$FLAGS = $FC_OVERWRITE + $FC_CREATEPATH
	Else
		$FLAGS = $FC_NOOVERWRITE + $FC_CREATEPATH
	EndIf

	; Get list of files
	LogMessage("Searching " & $sCopySourceFolder & " for " & $sCopyFileName, 5, "CopyFiles")
	$asCopyFiles = _FileListToArray($sCopySourceFolder, $sCopyFileName)

	; No files found
	If UBound($asCopyFiles) = 0 Then
		LogMessage("No files found.", 4, "CopyFiles")
		SetError(1)
		Return 0
	EndIf

	; Copy all files found
	For $iCurrentFile = 1 To UBound($asCopyFiles) - 1
		$sCurrentFile = $asCopyFiles[$iCurrentFile]
		If FileCopy($sCopySourceFolder & "\" & $sCurrentFile, $sCopyDestinationFolder & "\" & $sCurrentFile, $FLAGS) Then
			LogMessage("Successfully copied " & $sCurrentFile, 4, "CopyFiles")
		Else
			If Not @error Then
				; File was not copied but no error encountered: i.e. no overwrite
				LogMessage($sCurrentFile & " already exists, not overwritten.", 4, "CopyFiles")
			Else
				; File system error
				ThrowError("Failed to copy " & $sCurrentFile, 3, "CopyFiles", @error)
				SetError(2)
			EndIf
		EndIf
	Next ;$file

	; Exit
	Return (Not @error)

EndFunc

; Function to copy a file or set of files
Func MoveFiles($sMoveFileName, $sMoveSourceFolder, $sMoveDestinationFolder, $bOverwrite = False)
	LogMessage("Called MoveFiles($sMoveFileName = " & $sMoveFileName & ", $sMoveSourceFolder = " & $sMoveSourceFolder & ", $sMoveDestinationFolder = " & $sMoveDestinationFolder & ", $bOverwrite = " & $bOverwrite & ")", 5)

	Local $asMoveFiles
	Local $iCurrentFile
	Local $sCurrentFile
	Local $FLAGS

	; Should overwrite?
	If $bOverwrite Then
		$FLAGS = $FC_OVERWRITE + $FC_CREATEPATH
	Else
		$FLAGS = $FC_NOOVERWRITE + $FC_CREATEPATH
	EndIf

	; Get list of files
	LogMessage("Searching " & $sMoveSourceFolder & " for " & $sMoveFileName, 5, "MoveFiles")
	$asMoveFiles = _FileListToArray($sMoveSourceFolder, $sMoveFileName)

	; No files found
	If UBound($asMoveFiles) = 0 Then
		LogMessage("No files found.", 4, "MoveFiles")
		SetError(1)
		Return 0
	EndIf

	; Copy all files found
	For $iCurrentFile = 1 To UBound($asMoveFiles) - 1
		$sCurrentFile = $asMoveFiles[$iCurrentFile]
		If FileMove($sMoveSourceFolder & "\" & $sCurrentFile, $sMoveDestinationFolder & "\" & $sCurrentFile, $FLAGS) Then
			LogMessage("Successfully moved " & $sCurrentFile & " to " & $sMoveDestinationFolder, 4, "MoveFiles")
		Else
			If Not @error Then
				; File was not copied but no error encountered: i.e. no overwrite
				LogMessage($sCurrentFile & " already exists, not overwritten.", 4, "MoveFiles")
			Else
				; File system error
				ThrowError("Failed to move " & $sCurrentFile, 3, "MoveFiles", @error)
				SetError(2)
			EndIf
		EndIf
	Next ;$file

	; Exit
	Return (Not @error)

EndFunc

; Function to delete a file or set of files
Func DeleteFiles($sDeleteFileName, $sSearchFolder = @WorkingDir)
	LogMessage("Called DeleteFiles($sDeleteFileName = " & $sDeleteFileName & ", $sSearchFolder = " & $sSearchFolder & ")", 5)

	Local $asDeleteFiles
	Local $iCurrentFile
	Local $sCurrentFile
	Local $FLAGS

	; Get list of files
	LogMessage("Searching for " & $sDeleteFileName, 5, "DeleteFiles")
	$asDeleteFiles = _FileListToArray($sSearchFolder, $sDeleteFileName)

	; If no files found then exit function gracefully
	If UBound($asDeleteFiles) = 0 Then
		LogMessage("No files found.", 4, "DeleteFiles")
		SetError(0)
		Return 0
	EndIf

	; Delete all files found
	For $iCurrentFile = 1 To UBound($asDeleteFiles) - 1
		$sCurrentFile = $asDeleteFiles[$iCurrentFile]
		If FileDelete($sSearchFolder & "\" & $sCurrentFile) Then
			LogMessage("Successfully deleted " & $sCurrentFile, 4, "DeleteFiles")
		Else
			ThrowError("Failed to delete " & $sCurrentFile, 3, "DeleteFiles", @error)
			SetError(1)
		EndIf
	Next ; file

	; Exit
	Return (Not @error)

EndFunc

; Function to write to log
Func LogMessage($sMessageText, $iImportance = 3, $sFunctionName = "", $sLogFile = $g_sLogFile, $sWorkingDirectory = @WorkingDir)

	; Declarations
	Local $fMessage
	Local $sLogMessage = ""
	Local $sMarkdownMessage = ""
	Local $asResult

	; Errors get special handling
	If StringInStr($sMessageText, "ERROR") Then

		; Log message already created by the error handler
		$sLogMessage = $sMessageText
		$sMarkdownMessage = $sMessageText

	Else

		; Create message from function name and message text
		If $sFunctionName Then
			$sLogMessage &= "[" & $sFunctionName & "] " & $sMessageText
		Else
			$sLogMessage = $sMessageText
		EndIf

		; Start of the program gets special handling in Markdown version
		$asResult = StringRegExp($sMessageText, "Starting ([a-zA-Z]+LORASR)", 1)
		If UBound($asResult) > 0 Then
			; Add a header to the Markdown log file
			$sMarkdownMessage = @CRLF & "--------------------------------------------------------------------------------" & @CRLF & @CRLF
			$sMarkdownMessage &= "# " & $asResult[0] & @CRLF & @CRLF
			If $sFunctionName Then $sMarkdownMessage &= "[" & $sFunctionName & "] "
			$sMarkdownMessage &= $sMessageText & @CRLF & @CRLF
			$sMarkdownMessage &= "--------------------------------------------------------------------------------" & @CRLF & @CRLF
		Else
			; Importance level 1 and 2 messages get headings
			Switch $iImportance
				Case 1
					$sMarkdownMessage = @CRLF & "--------------------------------------------------------------------------------" & @CRLF & @CRLF
					$sMarkdownMessage &= "# " & Heading($sMessageText)
					If $sFunctionName Then $sMarkdownMessage &= @CRLF & @CRLF & "[" & $sFunctionName & "] " & $sMessageText
				Case 2
					$sMarkdownMessage = "## " & Heading($sMessageText)
					If $sFunctionName Then $sMarkdownMessage &= @CRLF & @CRLF & "[" & $sFunctionName & "] " & $sMessageText
				Case Else
					$sMarkdownMessage = $sLogMessage
			EndSwitch
			; Start of each run in a batch gets special Handling
			If StringLeft($sMessageText, 12) = "Starting run" Then $sMarkdownMessage = @CRLF & "--------------------------------------------------------------------------------" & @CRLF & @CRLF & $sMarkdownMessage
			; 'Tidying up leftover files' comes after all runs are complete
			If StringInStr($sMessageText, "leftover") Then $sMarkdownMessage = @CRLF & "--------------------------------------------------------------------------------" & @CRLF & @CRLF & $sMarkdownMessage
			; Add spaces after the message body to give a new line in Markdown
			$sMarkdownMessage &= "  "
		EndIf
	EndIf

	; Write to console (when running from development environment)
	If $iImportance <= $g_iConsoleVerbosity Then
		; Blank line to separate out important messages
		If ($iImportance <= 2) And ($g_iConsoleVerbosity > 2) Then ConsoleWrite(@CRLF)
		; Write message
		ConsoleWrite($sLogMessage & @CRLF)
	EndIf

	; Write to log
	If $iImportance <= $g_iLogFileVerbosity Then
		; Create new log file if it doesn't already exist
		If Not FileExists($sWorkingDirectory & "\" & $sLogFile) Then CreateLogFile($sLogFile, $sWorkingDirectory)
		; Blank line to separate out important messages
		If ($iImportance <= 2) And ($g_iLogFileVerbosity > 2) Then WriteToLogFile("", $sLogFile, $sWorkingDirectory)
		; Write message
		WriteToLogFile($sMarkdownMessage, $sLogFile, $sWorkingDirectory)
	EndIf

	; Message box
	If $iImportance <= $g_iMessageVerbosity Then
		Switch $iImportance
			Case 1
				$fMessage = $MB_SYSTEMMODAL + $MB_OK
			Case Else
				$fMessage = $MB_OK
		EndSwitch
		MsgBox($fMessage, $sFunctionName, $sMessageText)
	EndIf

	; Exit function
	Return (Not @error)

EndFunc

; Function to write to the log file
Func WriteToLogFile($sMessageText, $sLogFile = $g_sLogFile, $sWorkingDirectory = @WorkingDir)
	; Note there is no logging in this function, as that could create a recursive loop.

	; Declarations
	Local $hLogFile = 0

	; Open file
	$hLogFile = FileOpen($sWorkingDirectory & "\" & $sLogFile, $FO_APPEND)

	; Write text
	FileWriteLine($hLogFile, $sMessageText)

	; Close file
	FileClose($hLogFile)

	; Exit
	Return (Not @error)

EndFunc

; Function to create a new log file
Func CreateLogFile($sLogFile = $g_sLogFile, $sWorkingDirectory = @WorkingDir)
	; Note there is no logging in this function, as that could create a recursive loop.

	; Declarations
	Local $hLogFile = 0
	Local $tCurrentTime

	; If file exists, move it out of the way
	If FileExists($sWorkingDirectory & "\" & $sLogFile) Then FileMove($sWorkingDirectory & "\" & $sLogFile, $sWorkingDirectory & "\" & $sLogFile & ".old", $FC_OVERWRITE)

	; Create new file
	$hLogFile = FileOpen($sWorkingDirectory & "\" & $sLogFile, $FO_APPEND)

	; Write headers
	FileWriteLine($hLogFile, "--------------------------------------------------------------------------------")
	FileWriteLine($hLogFile, "")
	FileWriteLine($hLogFile, "# runLORASR log file")
	FileWriteLine($hLogFile, "")
	FileWriteLine($hLogFile, "Working folder: `" & $sWorkingDirectory & "`  ")
	$tCurrentTime = _Date_Time_GetLocalTime()
	FileWriteLine($hLogFile, "Run date/time:  " & _Date_Time_SystemTimeToDateTimeStr($tCurrentTime,1))
	FileWriteLine($hLogFile, "")
	FileWriteLine($hLogFile, "--------------------------------------------------------------------------------")
	FileWriteLine($hLogFile, "")
	FileWriteLine($hLogFile, "## Loading runLORASR libraries")
	FileWriteLine($hLogFile, "")

	; Close file
	FileClose($hLogFile)

	; Exit
	Return (Not @error)

EndFunc

; Function to handle errors
Func ThrowError($sErrorText = "", $iImportance = 3, $sFunctionName = "", $iErrorCode = 0, $sLogFile = $g_sLogFile, $sWorkingDirectory = @WorkingDir)
	; Note there is no logging in this function, as that could create a recursive loop.

	; Declarations
	Local $sErrorMessage = ""

	; Build error message from given information
	$sErrorMessage = @CRLF & "*** ERROR "
	If $sFunctionName Then $sErrorMessage &= "in " & $sFunctionName & " "
	$sErrorMessage &= "***  " & @CRLF
	If $sErrorText Then $sErrorMessage &= $sErrorText & "  " & @CRLF
	If $iErrorCode Then $sErrorMessage &= "Error code: " & String($iErrorCode) & "  " & @CRLF
	$sErrorMessage &= @CRLF

	; Send the message
	LogMessage($sErrorMessage, $iImportance, $sFunctionName, $sLogFile, $sWorkingDirectory)

	; Set error code
	If $iErrorCode Then
		SetError($iErrorCode)
	Else
		SetError(1)
	EndIf

	; Exit
	Return 1

EndFunc

; Function to make a heading from a message
Func Heading($sText)
	; No logging as this is part of the logging process

	; Trim any trailing dots
	While StringRight($sText, 1) = "."
		$sText = StringTrimRight($sText, 1)
	Wend

	; Keep just the first part before ":" or " - "
	If StringInStr($sText, ":") Then $sText = StringLeft($sText, StringInStr($sText, ":") - 1)
	If StringInStr($sText, " - ") Then $sText = StringLeft($sText, StringInStr($sText, " - ") - 1)

	; Lose the detail after connecting words
	If StringInStr($sText, " to ") Then $sText = StringLeft($sText, StringInStr($sText, " to ") - 1)
	If StringInStr($sText, " of ") Then $sText = StringLeft($sText, StringInStr($sText, " of ") - 1)
	If StringInStr($sText, " for ") Then $sText = StringLeft($sText, StringInStr($sText, " for ") - 1)

	; Return modified string
	Return $sText

EndFunc
