#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Batch
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.08.08
 Modified:       2017.08.11
 Version:        0.4.0.28

 Script Function:
	Work through a batch of input files and run LORASR for each one

#ce ----------------------------------------------------------------------------

#include-once
#include <File.au3>
#include <FileConstants.au3>
#include <MsgBoxConstants.au3>
#include "runLORASR.Functions.au3"
#include "runLORASR.Sweep.au3"
#include "runLORASR.Run.au3"
#include "runLORASR.Plots.au3"
#include "runLORASR.Tidy.au3"

LogMessage("Loaded runLORASR.Batch version 0.4.0.28", 3)

; Main function
Func BatchLORASR($sWorkingDirectory = @WorkingDir, $sProgramPath = "C:\Program Files (x86)\LORASR", $sSimulationProgram = "LORASR.exe", $sSweepFile = "Sweep.xlsx", $sTemplateFile = "Template.txt", $sResultsFile = "Batch results.csv", $sPlotFile = "Plots.xlsx", $sInputFolder = "Input", $sOutputFolder = "Output", $sRunFolder = "Runs", $sIncompleteFolder = "Incomplete", $bCleanup = True)
	LogMessage("Called BatchLORASR($sWorkingDirectory = " & $sWorkingDirectory & ", $sProgramPath = " & $sProgramPath & ", $sSimulationProgram = " & $sSimulationProgram & ", $sSweepFile = " & $sSweepFile & ", $sTemplateFile = " & $sTemplateFile & ", $sResultsFile = " & $sResultsFile & ", $sPlotFile = " & $sPlotFile & ", $sInputFolder = " & $sInputFolder & ", $sOutputFolder = " & $sOutputFolder & ", $sRunFolder = " & $sRunFolder & ", $sIncompleteFolder = " & $sIncompleteFolder & ", $bCleanup = " & $bCleanup & ")", 5)

	; Declarations
	Local $iResult = 0, $iCurrentInputFile = 0
	Local $bCreateResultsFile = True
	Local $sRun = "", $sSimulationProgramPath = "", $sStart = "", $sEnd = ""
	Local $asInputFiles
	Local $tStart, $tEnd

	; Switch to working directory
	FileChangeDir($sWorkingDirectory)
	If @error Then
		ThrowError("Could not access working directory", 1, "BatchLORASR", @error)
		SetError (1)
		Return 0
	EndIf

	; Copy input files to working directory
	LogMessage("Searching for input files...", 2, "BatchLORASR")
	$iResult = FindInputFiles($sWorkingDirectory, $sInputFolder, $sSweepFile, $sTemplateFile, $sPlotFile, $sProgramPath)
	If (Not $iResult) Or @error Then
		; Errors at this stage may not stop the batch from running, but should be noted.
		ThrowError("Could not find any input files. Batch cancelled.", 1, "BatchLORASR", @error)
		SetError(2)
		Return 0
	EndIf

	; Try to set up parameter sweep
	LogMessage("Checking for parameter sweep definition...", 2, "BatchLORASR")
	If RunSweepLORASR($sWorkingDirectory, $sSweepFile, $sTemplateFile, $sResultsFile, $sInputFolder) Then
		; If parameter sweep preparations were successful, the results file should already be created, so don't create it again.
		$bCreateResultsFile = False
	Else
		If @error Then
			; RunSweepLORASR returns an error only if the batch should be cancelled
			ThrowError("Batch cancelled.", 1, "BatchLORASR", @error)
			SetError(3)
			Return 0
		Else
			; No parameter sweep, so need to create a batch results file (below)
			$bCreateResultsFile = True
		EndIf
	EndIf

	; Create results file unless already created by sweep program
	If $bCreateResultsFile Then
		LogMessage("Creating batch results output file...", 2, "BatchLORASR")
		$iResult = CreateResultsFile(0, $sWorkingDirectory, $sResultsFile)
		If (Not $iResult) Or @error Then
			; Write failed
			If MsgBox($MB_OKCANCEL, "batchLORASR", "Error creating results file " & $sResultsFile & "." & @CRLF & "Do you want to continue?") = $IDCANCEL Then
				ThrowError("Batch cancelled.", 1, "BatchLORASR", @error)
				SetError(4)
				Return 0
			EndIf
		EndIf
	EndIf

	; Set up for first run
	LogMessage("Setting up simulation environment...", 2, "BatchLORASR")
	$sSimulationProgramPath = SetupLORASR($sWorkingDirectory, $sProgramPath, $sSimulationProgram)
	If (Not $sSimulationProgramPath) Or @error Then
		ThrowError("Could not set up simulation environment in folder " & $sWorkingDirectory & ". Batch cancelled.", 1, "BatchLORASR", @error)
		SetError(5)
		Return 0
	EndIf

	; Get list of input files in working directory
	LogMessage("Starting to process input files...", 2, "BatchLORASR")
	$asInputFiles = _FileListToArray($sWorkingDirectory, "*.in")
	If (UBound($asInputFiles) = 0) Or @error Then
		ThrowError("No input files found in " & $sWorkingDirectory & " nor in input subfolder '" & $sInputFolder & "'. Batch cancelled.", 1, "BatchLORASR", @error)
		SetError(6)
		Return 0
	EndIf

	; ---------------------------------------------------------------------------------
	; Run process for each input file sequentially
	For $iCurrentInputFile = 1 To UBound($asInputFiles) - 1

		; Define run name by stripping file extension ".in" from input file
		$sRun = StringTrimRight($asInputFiles[$iCurrentInputFile], 3)

		; Start the run
		LogMessage("Starting run " & $sRun & "...", 2, "BatchLORASR")
		$tStart = _Date_Time_GetLocalTime()
		$sStart = _Date_Time_SystemTimeToDateTimeStr($tStart, 1)
		LogMessage("Start time: " & $sStart, 4, "BatchLORASR")

		; Call the main run function
		$iResult = RunLORASR($sRun, $sWorkingDirectory, $sSimulationProgramPath, $sInputFolder)

		; Report result
		LogMessage("Run " & $sRun & " Result: " & $iResult, 2, "BatchLORASR")
		$tEnd = _Date_Time_GetLocalTime()
		$sEnd = _Date_Time_SystemTimeToDateTimeStr($tEnd, 1)
		LogMessage("Current time: " & $sEnd, 4, "BatchLORASR")

		; Try once more if failed
		If Not ($iResult = 1) Then
			LogMessage("Starting run " & $sRun & " again...", 2, "BatchLORASR")
			$iResult = RunLORASR($sRun, $sWorkingDirectory, $sSimulationProgramPath, $sInputFolder)
			LogMessage("Run " & $sRun & " Result: " & $iResult, 2, "BatchLORASR")
			$tEnd = _Date_Time_GetLocalTime()
			$sEnd = _Date_Time_SystemTimeToDateTimeStr($tEnd, 1)
			LogMessage("Current time: " & $sEnd, 4, "BatchLORASR")
			If Not ($iResult = 1) Then
				; Log failure and continue
				ThrowError("Run " & $sRun & " failed, attempting to continue with batch.", 2, "BatchLORASR", @error)
				TidyIncompleteRun($sRun, $sWorkingDirectory, $sIncompleteFolder)
				SetError(0)
				ContinueLoop
			EndIf
		EndIf

		; Save run results to output file
		LogMessage("Saving run " & $sRun & " to batch results output file...", 2, "BatchLORASR")
		$iResult = SaveResults($sRun, $sResultsFile, $sWorkingDirectory)
		If (Not $iResult) Or @error Then
			; Log failure and continue
			ThrowError("Error saving run " & $sRun & " to batch results output file.", 2, "BatchLORASR", @error)
			TidyIncompleteRun($sRun, $sWorkingDirectory, $sIncompleteFolder)
			SetError(0)
			ContinueLoop
		EndIf

		; Create plots spreadsheet
		LogMessage("Saving results of run " & $sRun & " to spreadsheet and creating plots...", 2, "BatchLORASR")
		$iResult = PlotLORASR($sWorkingDirectory, $sRun & ".xlsx", $sPlotFile, $sProgramPath)
		If (Not $iResult) Or @error Then
			; Log failure and continue
			ThrowError("Error saving run " & $sRun & " to plots spreadsheet.", 2, "BatchLORASR", @error)
			TidyIncompleteRun($sRun, $sWorkingDirectory, $sIncompleteFolder)
			SetError(0)
			ContinueLoop
		EndIf

		; Tidy up
		If $bCleanup Then
			LogMessage("Tidying up files for run " & $sRun & "...", 2, "BatchLORASR")
			$iResult = TidyCompletedRun($sRun, $sWorkingDirectory, $sInputFolder, $sOutputFolder, $sRunFolder)
			If (Not $iResult) Or @error Then
				; Log failure and continue
				ThrowError("Could not tidy up for run " & $sRun & ".", 3, "BatchLORASR", @error)
				SetError(0)
				ContinueLoop
			Else
				LogMessage("Tidying up complete.", 4, "BatchLORASR")
			EndIf
		EndIf

		; Report completion
		LogMessage("Run " & $sRun & " complete.", 2, "BatchLORASR")
		$tEnd = _Date_Time_GetLocalTime()
		$sEnd = _Date_Time_SystemTimeToDateTimeStr($tEnd, 1)
		LogMessage("End time: " & $sEnd, 4, "BatchLORASR")
		LogMessage("Run time: " & _DateDiff ("m", $sStart, $sEnd) & "m " & _DateDiff ("s", $sStart, $sEnd) & "s", 4, "BatchLORASR")

	Next
	; ---------------------------------------------------------------------------------

	; Tidy up
	If $bCleanup Then
		LogMessage("Tidying up leftover files...", 2, "BatchLORASR")
		$iResult = TidyBatchFiles($sWorkingDirectory, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder)
		If (Not $iResult) Or @error Then
			LogMessage("Finished tidying with some errors.", 3, "BatchLORASR")
			; Tidying errors can be ignored
			SetError(0)
		Else
			LogMessage("Finished tidying.", 3, "BatchLORASR")
		EndIf
	EndIf

	; Exit
	Return (Not @error)

EndFunc

; Function to get input files from separate input folder if present
Func FindInputFiles($sWorkingDirectory = @WorkingDir, $sInputFolder = "Input", $sSweepFile = "Sweep.xlsx", $sTemplateFile = "Template.txt", $sPlotFile = "Plots.xlsx", $sProgramPath = "C:\Program Files (x86)\LORASR")
	LogMessage("Called FindInputFilesFindInputFiles($sWorkingDirectory = " & $sWorkingDirectory & ", $sInputFolder = " & $sInputFolder & ", $sSweepFile = " & $sSweepFile & ", $sTemplateFile = " & $sTemplateFile & ", $sPlotFile = " & $sPlotFile & ", $sProgramPath = " & $sProgramPath & ")", 5)

	; Declarations
	Local $sFoundSweepFile = "", $sFoundTemplateFile = "", $sFoundInputFile = "", $sFoundPlotFile = ""

	; Look for sweep definition files
	LogMessage("Searching for sweep definition files", 3, "FindInputFiles")
	$sFoundSweepFile = FindFile($sSweepFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
	$sFoundTemplateFile = FindFile($sTemplateFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)

	; If there is no parameter sweep, look for standard input files
	If ($sFoundSweepFile And $sFoundTemplateFile) Then
		LogMessage("Found sweep definition files", 3, "FindInputFiles")
	Else
		LogMessage("Searching for LORASR input files", 3, "FindInputFiles")
		$sFoundInputFile = FindFile("*.in", $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
		If $sFoundInputFile Then
			LogMessage("Found LORASR input files", 3, "FindInputFiles")
		Else
			ThrowError("Could not find input files", 3, "FindInputFiles")
			SetError(1)
			Return 0
		EndIf
	EndIf

	; Copy master plot file
	LogMessage("Searching for master plot file", 3, "FindInputFiles")
	$sFoundPlotFile = FindFile($sPlotFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
	If $sFoundPlotFile Then
		LogMessage("Found plot file", 3, "FindInputFiles")
	Else
		; Main master should be in the program directory; doesn't need to be copied
		$sFoundPlotFile = FindFile($sPlotFile, $sWorkingDirectory, $sProgramPath, False)
		If $sFoundPlotFile Then
			LogMessage("Found master plot file", 3, "FindInputFiles")
		Else
			ThrowError("Could not find master plot file", 3, "FindInputFiles")
			SetError(2)
			Return 0
		EndIf
	EndIf

	; Exit
	Return (Not @error)

EndFunc

; Function to search for sweep parameter definition and call SweepLORASR if found
Func RunSweepLORASR($sWorkingDirectory = @WorkingDir, $sSweepFile = "Sweep.xlsx", $sTemplateFile = "Template.txt", $sResultsFile = "Batch results.csv", $sInputFolder = "Input")
	LogMessage("Called RunSweepLORASR($sWorkingDirectory = " & $sWorkingDirectory & ", $sSweepFile = " & $sSweepFile & ", $sTemplateFile = " & $sTemplateFile & ", $sResultsFile = " & $sResultsFile & ", $sInputFolder = " & $sInputFolder & ")", 5)

	If FileExists($sWorkingDirectory & "\" & $sSweepFile) Then
		If FileExists($sWorkingDirectory & "\" & $sTemplateFile) Then
			; Run sweep function
			$iResult = SweepLORASR($sWorkingDirectory, $sSweepFile, $sTemplateFile, $sResultsFile, $sInputFolder)
			; Check result
			If $iResult = 1 And (Not @error) Then
				; Sweep setup complete
				Return 1
			Else
				; Sweep error: ask user whether they want to continue
				If MsgBox($MB_OKCANCEL, "RunSweepLORASR", "Errors encountered while preparing parameter sweep." & @CRLF & "Do you still want to continue?") = $IDCANCEL Then
					; Exit
					ThrowError("User cancelled process - parameter sweep preparations incomplete.", 2, "RunSweepLORASR", @error)
					SetError(1)
					Return 0
				EndIf
			EndIf
		Else
			; Missing template: ask user whether they want to continue
			If MsgBox($MB_OKCANCEL, "RunSweepLORASR", "Found sweep definition file " & $sSweepFile & " but cannot find run template " & $sTemplateFile & "." & @CRLF & "Do you want to continue as a standard batch run instead?") = $IDCANCEL Then
				; Exit
				ThrowError("User cancelled process - missing template file.", 2, "RunSweepLORASR", @error)
				SetError(2)
				Return 0
			EndIf
		EndIf
	Else
		; No sweep definition: standard batch run
		LogMessage("No sweep definition found, continuing as standard batch run.", 3, "RunSweepLORASR")
		Return 0
	EndIf

	; Shouldn't get here
	Return 0

EndFunc

; Function to write run results to output file
Func SaveResults($sRun, $sResultsFile = "Batch results.csv", $sWorkingDirectory = @WorkingDir)
	LogMessage("Called SaveResults($sRun = " & $sRun & ", $sResultsFile = " & $sResultsFile & ", $sWorkingDirectory = " & $sWorkingDirectory & ")", 5)

	; Declarations
	Local $asOutputFiles, $asOutputData, $asRunDetails
	Local $sCurrentOutputFile = "", $sTransmission = "", $sRunDate = "", $sRunTime = ""
	Local $iCurrentLine = 0, $iRunDetails = 0, $iCurrentDetail = 0
	Local $hResultsFile = 0

	; Get list of output files for current run
	LogMessage("Searching for output files for run " & $sRun, 4, "SaveResults")
	$asOutputFiles = _FileListToArray($sWorkingDirectory, $sRun & "-*.out")
	If (UBound($asOutputFiles) = 0) Or @error Then
		ThrowError("Error getting list of output files for run " & $sRun, 2, "SaveResults", @error)
		SetError(1)
		Return 0
	EndIf

	; Read in output file
	$sCurrentOutputFile = $asOutputFiles[UBound($asOutputFiles)-1] ; reads most recent file, assuming files are sorted by date and time based on filename
	LogMessage("Reading output file " & $sCurrentOutputFile, 4, "SaveResults")
	$asOutputData = FileReadToArray($sCurrentOutputFile)
	If (UBound($asOutputData) = 0) Or @error Then
		ThrowError("Error loading data from output file " & $sCurrentOutputFile, 2, "SaveResults", @error)
		SetError(2)
		Return 0
	EndIf

	; Find transmission data
	LogMessage("Loading transmission results", 4, "SaveResults")
	$iCurrentLine = 0
	Do
		$iCurrentLine += 1
	Until StringLeft($asOutputData[$iCurrentLine],41) = "COMMON CORE-PART. OF INVESTIG. PLANES/%= "
	$sTransmission = StringTrimLeft($asOutputData[$iCurrentLine], 42)
	If (Not $sTransmission) Or @error Then
		ThrowError("Error reading transmission data from output file " & $sCurrentOutputFile, 2, "SaveResults", @error)
		SetError(3)
		Return 0
	EndIf

	; File details
	LogMessage("Loading run details", 4, "SaveResults")
	$asRunDetails = StringSplit($sCurrentOutputFile, "-")
	$iRunDetails = UBound($asRunDetails) - 1
	$sRunDate = $asRunDetails[$iRunDetails - 1]
	$sRunTime = StringTrimRight($asRunDetails[$iRunDetails],4)
	If (Not $sRunDate) Or (Not $sRunTime) Or @error Then
		ThrowError("Error reading date and time details from output file " & $sCurrentOutputFile, 2, "SaveResults", @error)
		SetError(4)
		Return 0
	EndIf

	; Open results file
	$hResultsFile = FileOpen($sResultsFile, $FO_APPEND)
	If (Not $hResultsFile) Or @error Then
		ThrowError("Error opening batch results file " & $sResultsFile, 2, "SaveResults", @error)
		SetError(5)
		Return 0
	EndIf

	; Write out data to CSV
	LogMessage("Writing results to file " & $sResultsFile, 3, "SaveResults")
	$iResult = FileWrite($hResultsFile, $sRunDate & ", " & $sRunTime & ", " & $sTransmission)
	If (Not $iResult) Or @error Then
		ThrowError("Error writing to batch results file " & $sResultsFile, 2, "SaveResults", @error)
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
			ThrowError("Error writing to batch results file " & $sResultsFile, 2, "SaveResults", @error)
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
		ThrowError("Error writing to batch results file " & $sResultsFile, 2, "SaveResults", @error)
		FileClose($hResultsFile)
		SetError(8)
		Return 0
	EndIf

	; Close results file
	$iResult = FileClose($hResultsFile)
	If (Not $iResult) Or @error Then
		ThrowError("Error closing batch results file " & $sResultsFile, 2, "SaveResults", @error)
		SetError(9)
		Return 0
	EndIf

	; Exit
	Return (Not @error)

EndFunc
