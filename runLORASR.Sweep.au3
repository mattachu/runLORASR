#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Sweep
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.08.07
 Modified:       2017.09.05
 Version:        0.4.2.2

 Script Function:
    Create a batch of input files from a batch definition file and a template

#ce ----------------------------------------------------------------------------

#include-once
#include "runLORASR.Functions.au3"
#include "runLORASR.Results.au3"
#include <Array.au3>
#include <Excel.au3>
#include <File.au3>
#include <FileConstants.au3>

; Code version
$g_sSweepVersion = "0.4.2.2"

; Main function
Func SweepLORASR($sWorkingDirectory = @WorkingDir, $sSweepFile = "Sweep.xlsx", $sTemplateFile = "Template.txt", $sResultsFile = "Batch results.csv", $sInputFolder = "Input")
    LogMessage("Called `SweepLORASR($sWorkingDirectory = " & $sWorkingDirectory & ", $sSweepFile = " & $sSweepFile & ", $sTemplateFile = " & $sTemplateFile & ", $sResultsFile = " & $sResultsFile & ", $sInputFolder = " & $sInputFolder & ")`", 5)

    ; Declarations
    Local $iResult = 0, $iParameter = 0
    Local $asParameters, $asValues

    ; Find sweep definition files
    LogMessage("Finding sweep definition files...", 3, "SweepLORASR")
    $iResult = FindSweepFiles($sWorkingDirectory, $sSweepFile, $sTemplateFile, $sInputFolder)
    If (Not $iResult) Or @error Then
        LogMessage("Sweep definition files not found.", 3, "SweepLORASR")
        SetError(1)
        Return 0
    EndIf

    ; Load parameter details
    LogMessage("Loading parameters from sweep definition...", 3, "SweepLORASR")
    $iResult = LoadSweepParameters($asParameters, $asValues, $sWorkingDirectory, $sSweepFile)
    If (Not $iResult) Or @error Then
        ThrowError("Error loading sweep parameters. Cannot build parameter sweep.", 3, "SweepLORASR", @error)
        SetError(2)
        Return 0
    EndIf

    ; Clear up existing input files in working directory
    LogMessage("Clearing up existing input files...", 3, "SweepLORASR")
    $iResult = DeleteInputFiles($sWorkingDirectory, $sTemplateFile)
    If (Not $iResult) Or @error Then
        ThrowError("Error while clearing up input files", 4, "SweepLORASR", @error)
        ; Try to continue
        SetError(0)
    EndIf

    ; Build input files for each set of parameter values
    LogMessage("Building input files for parameter sweep...", 3, "SweepLORASR")
    For $iParameter = 0 To UBound($asParameters,1) - 1
        LogMessage("Parameter " & String($iParameter + 1) & " (" & $asParameters[$iParameter][0] & ")", 4, "SweepLORASR")
        $iResult = SweepParameter($asParameters, $asValues, $iParameter, $sWorkingDirectory, $sTemplateFile)
        If (Not $iResult) Or @error Then
            ThrowError("Error building input files. Cannot build parameter sweep.", 3, "SweepLORASR", @error)
            SetError(3)
            Return 0
        EndIf
    Next

    ; Create results file and write headers
    LogMessage("Creating batch results output file...", 3, "SweepLORASR")
    $iResult = CreateResultsFile($asParameters, $sWorkingDirectory, $sResultsFile)
    If (Not $iResult) Or @error Then
        ThrowError("Error building batch results output file.", 4, "SweepLORASR", @error)
        SetError(4)
    EndIf

    ; Exit
    LogMessage("End of parameter sweep preparations.", 3, "SweepLORASR")
    Return (Not @error)

EndFunc

; Function to clear up existing input files in the working directory
Func DeleteInputFiles($sWorkingDirectory = @WorkingDir, $sTemplateFile = "Template.txt")
    LogMessage("Called `DeleteInputFiles($sWorkingDirectory = " & $sWorkingDirectory & ", $sTemplateFile = " & $sTemplateFile & ")`", 5)

    ; Declarations
    Local $bMoveTemplate = False

    ; Check that the template file doesn't have the extension "in"
    If StringCompare(StringRight($sTemplateFile, 3), "*.in") = 0 Then
        $bMoveTemplate = True
        FileMove($sWorkingDirectory & "\" & $sTemplateFile, $sWorkingDirectory & "\" & StringTrimRight($sTemplateFile,3) & ".bak", $FC_OVERWRITE)
    EndIf
    If @error Then
        ThrowError("Error with template file " & $sTemplateFile, 3, "DeleteInputFiles", @error)
        SetError(1)
        Return 0
    EndIf

    ; Delete all *.in files
    DeleteFiles("*.in", $sWorkingDirectory)
    If @error Then
        ThrowError("Error deleting input files", 3, "DeleteInputFiles", @error)
        SetError(2)
    EndIf

    ; Put back the template file if required
    If $bMoveTemplate Then FileMove($sWorkingDirectory & "\" & StringTrimRight($sTemplateFile,3) & ".bak", $sWorkingDirectory & "\" & $sTemplateFile, $FC_OVERWRITE)

    ; Exit
    Return (Not @error)

EndFunc

; Function to find sweep files and copy them to the working directory
Func FindSweepFiles($sWorkingDirectory = @WorkingDir, $sSweepFile = "Sweep.xlsx", $sTemplateFile = "Template.txt", $sInputFolder = "Input")
    LogMessage("Called `FindSweepFiles($sWorkingDirectory = " & $sWorkingDirectory & ", $sSweepFile = " & $sSweepFile & ", $sTemplateFile = " & $sTemplateFile & ", $sInputFolder = " & $sInputFolder & ")`", 5)

    ; Find sweep definition file
    If FindFile($sSweepFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True) Then
        LogMessage("Found sweep parameters spreadsheet `" & $sSweepFile & "`", 4, "FindSweepFiles")
    Else
        ThrowError("Sweep parameters spreadsheet `" & $sSweepFile & "` not found.", 3, "FindSweepFiles", @error)
        SetError(1)
        Return 0
    EndIf

    ; Find template file
    If FindFile($sSweepFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True) Then
        LogMessage("Found template file `" & $sTemplateFile & "`", 4, "FindSweepFiles")
    Else
        ThrowError("Template file `" & $sTemplateFile & "` not found.", 3, "FindSweepFiles", @error)
        SetError(2)
        Return 0
    EndIf

    ; Exit
    Return (Not @error)

EndFunc

; Function to load sweep parameters from spreadsheet
Func LoadSweepParameters(ByRef $asParameters, ByRef $asValues, $sWorkingDirectory = @WorkingDir, $sSweepFile = "Sweep.xlsx")
    LogMessage("Called `LoadSweepParameters(ByRef $asParameters, ByRef $asValues, $sWorkingDirectory = " & $sWorkingDirectory & ", $sSweepFile = " & $sSweepFile & ")`", 5)

    ; Declarations
    Local $oExcel, $oWorkbook

    ; Open Excel
    $oExcel = _Excel_Open(False)
    If @error Then
        ThrowError("Error creating the Excel application object.", 3, "LoadSweepParameters", @error)
        SetError(1)
        Return 0
    EndIf

    ; Open the spreadsheet
    LogMessage("Opening workbook `" & $sWorkingDirectory & "\" & $sSweepFile & "`", 3, "LoadSweepParameters")
    $oWorkbook = _Excel_BookOpen($oExcel, $sWorkingDirectory & "\" & $sSweepFile)
    If @error Then
        ThrowError("Error opening workbook `" & $sSweepFile & "`", 3, "LoadSweepParameters", @error)
        SetError(2)
        Return 0
    EndIf

    ; Read in parameter details from Excel file
    $nParameters = _Excel_RangeRead($oWorkbook, 1, "B1")
    $asParameters = _Excel_RangeRead($oWorkbook, 1, "A4:D" & $nParameters + 3)
    If @error Then
        ThrowError("Error reading parameters from workbook.", 3, "LoadSweepParameters", @error)
        SetError(3)
        Return 0
    EndIf
    LogMessage("Loaded " & $nParameters & " parameter definitions", 3, "LoadSweepParameters")

    ; Read in parameter values from Excel file
    $nValues = _ArrayMax($asParameters,1,-1,-1,1)
    If $nParameters = 1 Then
        ; Workaround to force $values to be a 2D array to avoid errors when only one parameter is swept
        $asValues = _Excel_RangeRead($oWorkbook, 2, "A2:" & Chr(Asc("A") + $nParameters) & $nvalues + 1)
    Else
        $asValues = _Excel_RangeRead($oWorkbook, 2, "A2:" & Chr(Asc("A") + $nParameters - 1) & $nvalues + 1)
    EndIf
    If @error Then
        ThrowError("Error reading parameter values from workbook.", 3, "LoadSweepParameters", @error)
        SetError(34)
        Return 0
    EndIf
    LogMessage("Loaded parameter values for all defined parameters", 3, "LoadSweepParameters")

    ; Close Excel
    _Excel_Close($oExcel)
    If @error Then
        ThrowError("Error closing Excel.", 3, "LoadSweepParameters", @error)
        SetError(5)
    EndIf

    ; Exit
    Return (Not @error)

EndFunc

; Function to sweep through a parameter and create input files for each value of that parameter
Func SweepParameter($asParameters, $asValues, $iParameter = 0, $sWorkingDirectory = @WorkingDir, $sTemplateFile = "Template.txt")
    LogMessage("Called `SweepParameter($asParameters, $asValues, $iParameter = " & $iParameter & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $sTemplateFile = " & $sTemplateFile & ")`", 5)

    ; Declarations
    Local $iValue = 0
    Local $asBaseFiles
    Local $sBaseFile = "", $sCurrentFile

    ; Build list of base files from which to build the next level of the parameter sweep
    If $iParameter = 0 Then
        ; For the first parameter, create the set of input files from the template
        $sBaseFile  = $sTemplateFile
        Dim $asBaseFiles[1] = [$sBaseFile]
    Else
        ; For subsequent parameters, copy the existing input files and expand the parameter sweep.
        $asBaseFiles = _FileListToArray($sWorkingDirectory, "*.in")
        ; The first entry is a count of files, not required
        _ArrayDelete($asBaseFiles, 0)
    EndIf
    If @error Then
        ThrowError("Error getting list of existing input files.", 3, "SweepParameter", @error)
        SetError(1)
        Return 0
    EndIf

    ; Loop through all values for this parameter
    For $iValue = 0 To $asParameters[$iParameter][1] - 1

        ; Loop through all existing base files
        For $sBaseFile In $asBaseFiles

            ; Build filenames
            If $iParameter = 0 Then
                ; For first parameter, the filename is just the value of the first parameter
                $sCurrentFile = $asValues[$iValue][$iParameter] & ".in"
            Else
                ; For subsequent parameters, we need to skip the original template and modify the existing sweep files
                If $sBaseFile = $sTemplateFile Then ContinueLoop
                ; Add the value of the current parameter to name of the files that have already been swept
                $sCurrentFile = StringTrimRight($sBaseFile, 3) & "-" & $asValues[$iValue][$iParameter] & ".in"
            EndIf
            If @error Then
                ThrowError("Error building input file name for value #" & $iValue & " of parameter #" & $iParameter & " (" & $asParameters[$iParameter][0] & " = " & $asValues[$iValue][$iParameter] & ")", 3, "SweepParameter", @error)
                ; Move to next file
                SetError(0)
                ContinueLoop
            EndIf

            ; Copy files
            FileCopy($sWorkingDirectory & "\" & $sBaseFile, $sWorkingDirectory & "\" & $sCurrentFile, $FC_OVERWRITE)
            If @error Then
                ThrowError("Error creating input file `" & $sCurrentFile & "` from base file `" & $sBaseFile & "`", 3, "SweepParameter", @error)
                ; Move to next file
                SetError(0)
                ContinueLoop
            EndIf

            ; Replace the dummy code for this current parameter with the value of the parameter
            _ReplaceStringInFile($sWorkingDirectory & "\" & $sCurrentFile, $asParameters[$iParameter][3], $asValues[$iValue][$iParameter], $STR_CASESENSE)
            If @error Then
                ThrowError("Error modifying input file `" & $sCurrentFile & "`", 3, "SweepParameter", @error)
                ; Move to next file
                SetError(0)
                ContinueLoop
            EndIf

            ; Report success
            LogMessage("Created input file `" & $sCurrentFile & "`", 3, "SweepParameter")

        Next ; $asBaseFiles

    Next ; $iValue

    ; At the end of the value loop, delete the base files for previous parameters
    If $iParameter > 0 Then
        For $sBaseFile In $asBaseFiles
            FileDelete($sWorkingDirectory & "\" & $sBaseFile)
            If @error Then
                ThrowError("Error deleting base file `" & $sBaseFile & "`", 3, "SweepParameter", @error)
                ; Move to next file
                SetError(0)
                ContinueLoop
            EndIf
            LogMessage("Deleted base file `" & $sBaseFile & "`", 3, "SweepParameter")
        Next ; $sBaseFile
    EndIf

    ; Exit
    Return (Not @error)

EndFunc
