#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Tidy
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.13
 Modified:       2017.09.06
 Version:        0.4.3.2

 Script Function:
    Tidy up files from a batch run of LORASR

#ce ----------------------------------------------------------------------------

#include-once
#include "runLORASR.Functions.au3"
#include "runLORASR.Progress.au3"

; Code version
$g_sTidyVersion = "0.4.3.2"

; Tidy up files from an incomplete run
Func TidyIncompleteRun($sRun, $sWorkingDirectory = @WorkingDir, $sIncompleteFolder = "Incomplete")
    LogMessage("Called `TidyIncompleteRun($sRun = " & $sRun & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $sIncompleteFolder = " & $sIncompleteFolder & ")`", 5)

    ; Move input file to incomplete folder
    If FileExists($sWorkingDirectory & "\" & $sRun & ".in") Then MoveFiles($sRun & ".in", $sWorkingDirectory, $sWorkingDirectory & "\" & $sIncompleteFolder, True)
    ; Move any output files to incomplete folder as well
    If FileExists($sWorkingDirectory & "\" & $sRun & "*.out") Then MoveFiles($sRun & "*.out", $sWorkingDirectory, $sWorkingDirectory & "\" & $sIncompleteFolder, True)
    If FileExists($sWorkingDirectory & "\" & $sRun & ".xlsx") Then MoveFiles($sRun & ".xlsx", $sWorkingDirectory, $sWorkingDirectory & "\" & $sIncompleteFolder, True)

    ; Exit
    Return (Not @error)

EndFunc

; Tidy up files from a completed run
Func TidyCompletedRun($sRun, $sWorkingDirectory = @WorkingDir, $sInputFolder = "Input", $sOutputFolder = "Output", $sRunFolder = "Runs")
    LogMessage("Called `TidyCompletedRun($sRun = " & $sRun & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $sInputFolder = " & $sInputFolder & ", $sOutputFolder = " & $sOutputFolder & ", $sRunFolder = " & $sRunFolder & ")`", 5)

    ; Move input file to input folder
    If FileExists($sWorkingDirectory & "\" & $sRun & ".in") Then MoveFiles($sRun & ".in", $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
    ; Move simulation output files to run folder
    If FileExists($sWorkingDirectory & "\" & $sRun & "*.out") Then MoveFiles($sRun & "*.out", $sWorkingDirectory, $sWorkingDirectory & "\" & $sRunFolder, True)
    ; Move plot file to output folder
    If FileExists($sWorkingDirectory & "\" & $sRun & ".xlsx") Then MoveFiles($sRun & ".xlsx", $sWorkingDirectory, $sWorkingDirectory & "\" & $sOutputFolder, True)

    ; Exit
    Return (Not @error)

EndFunc

; Test whether a particular run was completed or not, and tidy up accordingly
Func TidyRunFiles($sRun, $sWorkingDirectory = @WorkingDir, $sInputFolder = "Input", $sIncompleteFolder = "Incomplete", $sOutputFolder = "Output", $sRunFolder = "Runs")
    LogMessage("Called `TidyRunFiles($sRun = " & $sRun & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $sInputFolder = " & $sInputFolder & ", $sIncompleteFolder = " & $sIncompleteFolder & ", $sOutputFolder = " & $sOutputFolder & ", $sRunFolder = " & $sRunFolder & ")`", 5)

    ; Declarations
    Local $asOutputFiles
    Local $iResult = 0
    Local $bIncomplete = False

    ; Find matching output files for the current run
    $asOutputFiles = _FileListToArray($sWorkingDirectory, $sRun & "-*.out")
    If @error Then
        If @error = 4 Then
            ; No files found
            $bIncomplete = True
        Else
            ; Error in filesystem operation
            ThrowError("Could not tidy up run " & $sRun, 3, "TidyRunFiles", @error)
            SetError(1)
            Return 0
        EndIf
    EndIf

    ; If no output file then there was a problem
    If $bIncomplete Then
        ; Incomplete run
        LogMessage("Run " & $sRun & " incomplete. Moving to incomplete folder.", 4, "TidyRunFiles")
        $iResult = TidyIncompleteRun($sRun, $sWorkingDirectory, $sIncompleteFolder)
        If (Not $iResult) Or @error Then
            ThrowError("Could not tidy up incomplete run " & $sRun, 3, "TidyRunFiles", @error)
            SetError(2)
            Return 0
        EndIf
    Else
        ; Completed run
        LogMessage("Run " & $sRun & " seems to have completed ok. Moving to standard folders.", 4, "TidyRunFiles")
        $iResult = TidyCompletedRun($sRun, $sWorkingDirectory, $sInputFolder, $sOutputFolder, $sRunFolder)
        If @error Then
            ThrowError("Could not tidy up completed run " & $sRun, 3, "TidyRunFiles", @error)
            SetError(3)
            Return 0
        EndIf
    EndIf

    ; Exit
    Return (Not @error)

EndFunc

; Work through all input files, test whether completed or not, and tidy up accordingly
Func TidyAllRunFiles($sWorkingDirectory = @WorkingDir, $sInputFolder = "Input", $sIncompleteFolder = "Incomplete", $sOutputFolder = "Output", $sRunFolder = "Runs")
    LogMessage("Called `TidyAllRunFiles($sWorkingDirectory = " & $sWorkingDirectory & ", $sInputFolder = " & $sInputFolder & ", $sIncompleteFolder = " & $sIncompleteFolder & ", $sOutputFolder = " & $sOutputFolder & ", $sRunFolder = " & $sRunFolder & ")`", 5)

    ; Declarations
    Local $sRun = ""
    Local $asInputFiles
    Local $iInputFiles = 0, $iCurrentInputFile = 0, $iResult = 0

    ; Get list of input files in working directory
    $asInputFiles = _FileListToArray(@WorkingDir, "*.in")
    If @error Then
        If @error = 4 Then
            ; No files found
            LogMessage("No input files found.", 4, "TidyAllRunFiles")
            SetError(0)
            Return 1
        Else
            ; Some other error
            ThrowError("Could not tidy up run files.", 3, "TidyAllRunFiles", @error)
            SetError(1)
            Return 0
        EndIf
    EndIf

    ; Check that there is at least one input file
    If $asInputFiles = 0 Then
        ; If no input files found, there's nothing to do.
        LogMessage("No input files found.", 4, "TidyAllRunFiles")
        SetError(0)
        Return 1
    Else
        ; Run through each input file sequentially
        $iInputFiles = UBound($asInputFiles) - 1
        For $iCurrentInputFile = 1 To $iInputFiles

            ; Strip file extension ".in" from run name
            $sRun = StringTrimRight($asInputFiles[$iCurrentInputFile], 3)
            If @error Then
                ThrowError("Could not find run name.", 3, "TidyAllRunFiles", @error)
                SetError(2)
                ContinueLoop
            EndIf

            ; Update progress meters
            UpdateProgress($g_sProgressType, Round($iCurrentInputFile/($iInputFiles + 2) * 100), "Run " & $iCurrentInputFile & " of " & $iInputFiles & ": " & $sRun)

            ; Tidy up files for this run
            $iResult = TidyRunFiles($sRun, $sWorkingDirectory, $sInputFolder, $sIncompleteFolder)
            If @error Then
                ThrowError("Could not tidy up for run " & $sRun, 3, "TidyAllRunFiles", @error)
                SetError(3)
                ContinueLoop
            EndIf

        Next
    EndIf

    ; Exit
    LogMessage("Finished checking run files.", 3, "TidyAllRunFiles")
    Return (Not @error)

EndFunc

; Tidy up all files used or created by the batch process
Func TidyBatchFiles($sWorkingDirectory = @WorkingDir, $sSimulationProgram = "LORASR.exe", $sSweepFile = "Sweep.xls", $sTemplateFile = "Template.txt", $sPlotFile = "Plots.xlsx", $sInputFolder = "Input", $sOutputFolder = "Output", $sRunFolder = "Runs")
    LogMessage("Called `TidyBatchFiles($sWorkingDirectory = " & $sWorkingDirectory & ", $sSimulationProgram = " & $sSimulationProgram & ", $sSweepFile = " & $sSweepFile & ", $sTemplateFile = " & $sTemplateFile & ", $sPlotFile = " & $sPlotFile & ", $sInputFolder = " & $sInputFolder & ", $sOutputFolder = " & $sOutputFolder & ", $sRunFolder = " & $sRunFolder & ")`", 5)

    ; Move input files to input directory
    LogMessage("Clearing up any input files", 3, "TidyBatchFiles")
    If FileExists($sWorkingDirectory & "\*.in") Then MoveFiles("*.in", $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
    ;If FileExists($sWorkingDirectory & "\" & $sSweepFile) Then MoveFiles($sSweepFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True) ;--- don't clear up the sweep parameters definition files, as these define the batch
    ;If FileExists($sWorkingDirectory & "\" & $sTemplateFile) Then MoveFiles($sTemplateFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
    If FileExists($sWorkingDirectory & "\" & $sPlotFile) Then MoveFiles($sPlotFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)

    ; Move run files to run directory
    LogMessage("Clearing up any run files", 3, "TidyBatchFiles")
    If FileExists($sWorkingDirectory & "\*.out") Then MoveFiles("*.out", $sWorkingDirectory, $sWorkingDirectory & "\" & $sRunFolder, True)

    ; Move output files to output directory
    LogMessage("Clearing up any output files", 3, "TidyBatchFiles")
    If FileExists($sWorkingDirectory & "\*.xlsx") Then MoveFiles("*.xlsx", $sWorkingDirectory, $sWorkingDirectory & "\" & $sOutputFolder, True)
    If FileExists($sWorkingDirectory & "\*.xls") Then MoveFiles("*.xls", $sWorkingDirectory, $sWorkingDirectory & "\" & $sOutputFolder, True)
    ;If FileExists($sWorkingDirectory & "\*.csv") Then MoveFiles("*.csv", $sWorkingDirectory, $sWorkingDirectory & "\" & $sOutputFolder, True) ;--- don't clear up the Batch results.csv file as this is a summary of the batch
    If FileExists($sWorkingDirectory & "\" & $sOutputFolder & "\" & $sSweepFile) Then MoveFiles($sSweepFile, $sWorkingDirectory & "\" & $sOutputFolder, $sWorkingDirectory, False) ; exclude sweep file from tidying (see above)

    ; Delete LORASR simulation files
    LogMessage("Clearing up LORASR simulation files", 3, "TidyBatchFiles")
    TidySimulationFiles($sWorkingDirectory)

    ; Delete LORASR executable
    LogMessage("Clearing up LORASR executable", 3, "TidyBatchFiles")
    If FileExists($sWorkingDirectory & "\" & $sSimulationProgram) Then DeleteFiles($sSimulationProgram, $sWorkingDirectory)

    ; Delete any old file versions
    LogMessage("Clearing up any old unneeded files", 3, "TidyBatchFiles")
    If FileExists($sWorkingDirectory & "\*.old") Then DeleteFiles("*.old", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\*.log") Or FileExists($sWorkingDirectory & "\*.log.md") Then TidyLogFiles($sWorkingDirectory, $sRunFolder)

    ; Exit
    Return (Not @error)

EndFunc

; Tidy up the files produced directly by the LORASR simulation code
Func TidySimulationFiles($sWorkingDirectory = @WorkingDir)
    LogMessage("Called `TidySimulationFiles($sWorkingDirectory = " & $sWorkingDirectory & ")`", 5)

    ; Delete LORASR data files
    If FileExists($sWorkingDirectory & "\bucent") Then DeleteFiles("bucent", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\DD") Then DeleteFiles("DD", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\emival") Then DeleteFiles("emival", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\fort.*") Then DeleteFiles("fort.*", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\interout") Then DeleteFiles("interout", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\long*") Then DeleteFiles("long*", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\phadv") Then DeleteFiles("phadv", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\rmsgrow") Then DeleteFiles("rmsgrow", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\transenv") Then DeleteFiles("transenv", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\xclu*") Then DeleteFiles("xclu*", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\yclu*") Then DeleteFiles("yclu*", $sWorkingDirectory)
    If FileExists($sWorkingDirectory & "\zclu*") Then DeleteFiles("zclu*", $sWorkingDirectory)

    ; Exit
    Return (Not @error)

EndFunc

Func TidyLogFiles($sWorkingDirectory = @WorkingDir, $sRunFolder = "Runs")
    LogMessage("Called `TidyLogFiles($sWorkingDirectory = " & $sWorkingDirectory & ")`", 5)

    ; Declarations
    Local $asLogFiles
    Local $iCurrentFile

    ; Get list of log files
    $asLogFiles = _FileListToArray(@WorkingDir, "*.log*")

    ; Loop through log files
    For $iCurrentFile = 1 To UBound($asLogFiles) - 1
        ; All files except the current log file get moved to the run folder
        If StringCompare($asLogFiles[$iCurrentFile], $g_sLogFile) Then MoveFiles($asLogFiles[$iCurrentFile], $sWorkingDirectory, $sWorkingDirectory & "\" & $sRunFolder, True)
    Next

    Return (Not @error)

EndFunc
