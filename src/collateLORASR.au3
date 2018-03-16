#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Work through a batch of input files and run LORASR for each one
#AutoIt3Wrapper_Res_Fileversion=0.4.5.0
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         collateLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.08.25
 Modified:       2018.03.16
 Version:        0.4.5.0

 Script Function:
    Collate transmission results from a set of LORASR output files

#ce ----------------------------------------------------------------------------

; Load libraries
#include "runLORASR.Functions.au3"
#include "runLORASR.Results.au3"

; Program version
Global CONST $g_sProgramName = "collateLORASR"
Global CONST $g_sProgramVersion = "0.4.5.0"

; Declarations
Local $iResult = 0
Local $sNewFolder = ""

; Create log file
$g_sLogFile = DateTimeFileName($g_sProgramName, ".log.md")
CreateLogFile($g_sLogFile, @WorkingDir)
LogVersions($g_sProgramName, $g_sProgramVersion)

; Get global settings
LogMessage("Loading global settings...", 2, "collateLORASR")
Local $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder
Local $bCleanup
$iResult = GetSettings(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)
If (Not $iResult) Or @error Then
    ThrowError("Error loading global settings", 1, "collateLORASR")
    Exit 2
EndIf

; Get working directory if run from program directory
LogMessage("Checking current directory...",2,  "collateLORASR")
If @WorkingDir = $sProgramPath Then
    $sNewFolder = FileSelectFolder("Please select the working directory that holds the LORASR output files.", "")
    If @error Then
        ThrowError("Could not select working directory", 1, "collateLORASR", @error)
        Exit 3
    EndIf
    FileChangeDir($sNewFolder)
    If @error Then
        ThrowError("Could not access working directory", 1, "collateLORASR", @error)
        Exit 4
    EndIf
EndIf
LogMessage("Results folder: `" & @WorkingDir & "`", 3,  "collateLORASR")

; Call the main run function
LogMessage("Starting collation batch...", 2,  "collateLORASR")
$iResult = SaveAllResults(@WorkingDir, $sResultsFile, $sInputFolder, $sRunFolder)

; Exit program
If ($iResult = 1) And (Not @error) Then
    LogMessage("Collation complete.", 1, "collateLORASR")
    Exit 0
Else
    LogMessage("Collation incomplete.", 1, "collateLORASR")
    Exit 1
EndIf
