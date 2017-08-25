#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Results
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.08.25
 Modified:       2017.08.25
 Version:        0.4.0.15

 Script Function:
	Extract transmission results from LORASR output files

#ce ----------------------------------------------------------------------------

#include-once
#include <Array.au3>
#include <File.au3>
#include <FileConstants.au3>
#include "runLORASR.Functions.au3"

LogMessage("Loaded runLORASR.Results version 0.4.0.15", 3)

; Function to loop through all output files and save results
Func SaveAllResults($sWorkingDirectory = @WorkingDir, $sResultsFile = "Batch results.csv", $sRunFolder = "Runs")
	LogMessage("Called SaveAllResults($sWorkingDirectory = " & $sWorkingDirectory & ", $sResultsFile = " & $sResultsFile & ", $sRunFolder = " & $sRunFolder & ")", 5)

	; Declarations
	Local $asOutputFiles, $asRuns
	Local $iCurrentRun = 0, $iResult = 0
	Local $sRun = ""

	; Get list of all output files
	LogMessage("Searching for output files", 4, "SaveAllResults")
	$asOutputFiles = _FileListToArray($sWorkingDirectory, "*.out")
	If (UBound($asOutputFiles) = 0) Or @error Then
		; If no files found, check for separate run folder
		If FileExists($sWorkingDirectory & "\" & $sRunFolder) Then
			LogMessage("Searching for output files in subfolder " & $sRunFolder, 4, "SaveAllResults")
			$sWorkingDirectory &= "\" & $sRunFolder
			$asOutputFiles = _FileListToArray($sWorkingDirectory, "*.out")
		EndIf
	EndIf
	If (UBound($asOutputFiles) = 0) Or @error Then
		; If still no files found, exit with error
		ThrowError("Error getting list of output files", 2, "SaveAllResults", @error)
		SetError(1)
		Return 0
	EndIf

	; Get list of unique runs (delete the 'count' row, remove the date, time and file extension, then return only unique rows)
	LogMessage("Compiling run list", 4, "SaveAllResults")
	$asRuns = $asOutputFiles
	_ArrayDelete($asRuns, 0)
	For $iCurrentRun = 0 To UBound($asRuns) - 1
		$asRuns[$iCurrentRun] = StringLeft($asRuns[$iCurrentRun], StringLen($asRuns[$iCurrentRun]) - 20)
	Next ; $iCurrentRun
	$asRuns = _ArrayUnique($asRuns)
	If (UBound($asRuns) = 0) Or @error Then
		ThrowError("Error compiling run list from output files", 2, "SaveAllResults", @error)
		SetError(2)
		Return 0
	EndIf

	; Create the results file
	LogMessage("Creating output file", 4, "SaveAllResults")
	$iResult = CreateResultsFile(0, $sWorkingDirectory, $sResultsFile)
	If (Not $iResult) Or @error Then
		ThrowError("Error creating results file", 2, "SaveAllResults", @error)
		SetError(3)
		Return 0
	EndIf

	; Loop through each run and save results to file
	LogMessage("Processing run results...", 4, "SaveAllResults")
	For $iCurrentRun = 1 To UBound($asRuns) - 1
		$sRun = $asRuns[$iCurrentRun]
		$iResult = SaveRunResults($sRun, $sWorkingDirectory, $sResultsFile)
		If (Not $iResult) Or @error Then
			LogMessage("Error saving results for run " & $sRun & ". Attempting to continue.", 3, "SaveAllResults")
			SetError(0)
			ContinueLoop
		EndIf
	Next ; $iCurrentRun

	; Log completion
	LogMessage("Batch complete.", 4, "SaveAllResults")

	; Exit
	Return (Not @error)

EndFunc

; Function to create results file to store the transmission results for each set of values in the parameter sweep
Func CreateResultsFile($asParameters, $sWorkingDirectory = @WorkingDir, $sResultsFile = "Batch results.csv")
	If $asParameters Then
		LogMessage("Called CreateResultsFile($asParameters = <array>, $sWorkingDirectory = " & $sWorkingDirectory & ", $sResultsFile = " & $sResultsFile & ")", 5)
	Else
		LogMessage("Called CreateResultsFile($asParameters = 0, $sWorkingDirectory = " & $sWorkingDirectory & ", $sResultsFile = " & $sResultsFile & ")", 5)
	EndIf

	; Declarations
	Local $hResultsFile = 0
	Local $sHeader = ""
	Local $iParameter = 0

	; Backup existing results file
	If FileExists($sWorkingDirectory & "\" & $sResultsFile) Then
		FileMove($sWorkingDirectory & "\" & $sResultsFile, $sWorkingDirectory & "\" & $sResultsFile & ".old", $FC_OVERWRITE)
		LogMessage("Backed up old results file to " & $sWorkingDirectory & "\" & $sResultsFile & ".old", 4, "CreateResultsFile")
	EndIf

	; Create new results file
	$hResultsFile = FileOpen($sWorkingDirectory & "\" & $sResultsFile, $FO_OVERWRITE + $FO_CREATEPATH)
	If @error Then
		ThrowError("Error creating results file.", 3, "CreateResultsFile", @error)
		FileClose($hResultsFile)
		SetError(1)
		Return 0
	EndIf
	LogMessage("Created new results file " & $sWorkingDirectory & "\" & $sResultsFile, 3, "CreateResultsFile")

	; Basic headers
	$sHeader = "Run date, Run time, Transmission (%)"

	; Parameter sweep or standard batch run?
	If UBound($asParameters) = 0 Then
		; If called with $asParameters set to zero, just write a generic header line
		$sHeader = $sHeader & ", Run name/details"
	Else
		; Parameter headers
		For $iParameter = 0 To UBound($asParameters,1) - 1
			$sHeader = $sHeader & ", " & $asParameters[$iParameter][0] & " (" & $asParameters[$iParameter][2] & ")"
		Next ; $iParameter
	EndIf

	; Write headers
	FileWriteLine($hResultsFile, $sHeader)
	If @error Then
		ThrowError("Error writing to results file.", 3, "CreateResultsFile", @error)
		SetError(2)
	EndIf

	; Close file
	FileClose($hResultsFile)
	If @error Then
		ThrowError("Error closing results file.", 4, "CreateResultsFile", @error)
		SetError(3)
		Return 0
	EndIf

	; Exit
	Return (Not @error)

EndFunc

; Function to extract and save results for a single run
Func SaveRunResults($sRun, $sWorkingDirectory = @WorkingDir, $sResultsFile = "Batch results.csv")
	LogMessage("Called SaveRunResults($sRun = " & $sRun & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $sResultsFile = " & $sResultsFile & ")", 5)

	; Declarations
	Local $asOutputFiles, $asOutputData, $asRunDetails
	Local $sCurrentOutputFile = "", $sTransmission = "", $sRunDate = "", $sRunTime = ""
	Local $iCurrentLine = 0, $iRunDetails = 0, $iCurrentDetail = 0
	Local $hResultsFile = 0

	; Get list of output files for current run
	LogMessage("Searching for output files for run " & $sRun, 4, "SaveRunResults")
	$asOutputFiles = _FileListToArray($sWorkingDirectory, $sRun & "-*.out")
	If (UBound($asOutputFiles) = 0) Or @error Then
		ThrowError("Error getting list of output files for run " & $sRun, 2, "SaveRunResults", @error)
		SetError(1)
		Return 0
	EndIf

	; Read in output file
	$sCurrentOutputFile = $asOutputFiles[UBound($asOutputFiles)-1] ; reads most recent file, assuming files are sorted by date and time based on filename
	LogMessage("Reading output file " & $sCurrentOutputFile, 4, "SaveRunResults")
	$asOutputData = FileReadToArray($sWorkingDirectory & "\" & $sCurrentOutputFile)
	If (UBound($asOutputData) = 0) Or @error Then
		ThrowError("Error loading data from output file " & $sCurrentOutputFile, 2, "SaveRunResults", @error)
		SetError(2)
		Return 0
	EndIf

	; Find transmission data
	LogMessage("Loading transmission results", 4, "SaveRunResults")
	$iCurrentLine = 0
	Do
		$iCurrentLine += 1
	Until StringLeft($asOutputData[$iCurrentLine],41) = "COMMON CORE-PART. OF INVESTIG. PLANES/%= "
	$sTransmission = StringTrimLeft($asOutputData[$iCurrentLine], 42)
	If (Not $sTransmission) Or @error Then
		ThrowError("Error reading transmission data from output file " & $sCurrentOutputFile, 2, "SaveRunResults", @error)
		SetError(3)
		Return 0
	EndIf

	; File details
	LogMessage("Loading run details", 4, "SaveRunResults")
	$asRunDetails = StringSplit($sCurrentOutputFile, "-")
	$iRunDetails = UBound($asRunDetails) - 1
	$sRunDate = $asRunDetails[$iRunDetails - 1]
	$sRunTime = StringTrimRight($asRunDetails[$iRunDetails],4)
	If (Not $sRunDate) Or (Not $sRunTime) Or @error Then
		ThrowError("Error reading run details from output file " & $sCurrentOutputFile, 2, "SaveRunResults", @error)
		SetError(4)
		Return 0
	EndIf

	; Open results file
	$hResultsFile = FileOpen($sResultsFile, $FO_APPEND)
	If (Not $hResultsFile) Or @error Then
		ThrowError("Error opening batch results file " & $sResultsFile, 2, "SaveRunResults", @error)
		SetError(5)
		Return 0
	EndIf

	; Write out data to CSV
	LogMessage("Writing results to file " & $sResultsFile, 3, "SaveRunResults")
	$iResult = FileWrite($hResultsFile, $sRunDate & ", " & $sRunTime & ", " & $sTransmission)
	If (Not $iResult) Or @error Then
		ThrowError("Error writing to batch results file " & $sResultsFile, 2, "SaveRunResults", @error)
		FileWrite($hResultsFile, @CRLF)
		FileClose($hResultsFile)
		SetError(6)
		Return 0
	EndIf

	; Write out additional run details based on filename
	$iCurrentDetail = 1
	While $iCurrentDetail < $iRunDetails - 1
		$iResult = FileWrite($hResultsFile, ", " & $asRunDetails[$iCurrentDetail])
		If (Not $iResult) Or @error Then
			ThrowError("Error writing to batch results file " & $sResultsFile, 2, "SaveRunResults", @error)
			FileWrite($hResultsFile, @CRLF)
			FileClose($hResultsFile)
			SetError(7)
			Return 0
		EndIf
		$iCurrentDetail += 1
	WEnd

	; Next line for next run
	$iResult = FileWrite($hResultsFile, @CRLF)
	If (Not $iResult) Or @error Then
		ThrowError("Error writing to batch results file " & $sResultsFile, 2, "SaveRunResults", @error)
		FileClose($hResultsFile)
		SetError(8)
		Return 0
	EndIf

	; Close results file
	$iResult = FileClose($hResultsFile)
	If (Not $iResult) Or @error Then
		ThrowError("Error closing batch results file " & $sResultsFile, 2, "SaveRunResults", @error)
		SetError(9)
		Return 0
	EndIf

	; Exit
	Return (Not @error)

EndFunc
