#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Work through a batch of input files and run LORASR for each one
#AutoIt3Wrapper_Res_Fileversion=0.4.2.5
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         batchLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.04
 Modified:       2017.09.05
 Version:        0.4.2.5

 Script Function:
    Work through a batch of input files and run LORASR for each one

#ce ----------------------------------------------------------------------------

; Load libraries
#include "runLORASR.Functions.au3"
#include "runLORASR.Batch.au3"

; Program version
Global CONST $g_sProgramName = "batchLORASR"
Global CONST $g_sProgramVersion = "0.4.2.5"

; Declarations
Local $iResult = 0
Local $sNewFolder = ""

; Create log file
$g_sLogFile = DateTimeFileName($g_sProgramName, ".log.md")
CreateLogFile($g_sLogFile, @WorkingDir)
LogVersions($g_sProgramName, $g_sProgramVersion)

; Get global settings
LogMessage("Loading global settings...", 2, "batchLORASR")
Local $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder
Local $bCleanup
$iResult = GetSettings(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)
If (Not $iResult) Or @error Then
    ThrowError("Error loading global settings", 1, "batchLORASR")
    Exit 2
EndIf

; Get working directory if run from program directory
LogMessage("Checking current directory...",2,  "batchLORASR")
If @WorkingDir = $sProgramPath Then
    $sNewFolder = FileSelectFolder("Please select the working directory in which to run LORASR.", "")
    If @error Then
        ThrowError("Could not select working directory", 1, "batchLORASR", @error)
        Exit 3
    EndIf
    FileChangeDir($sNewFolder)
    If @error Then
        ThrowError("Could not access working directory", 1, "batchLORASR", @error)
        Exit 4
    EndIf
EndIf
LogMessage("Running batch in working folder `" & @WorkingDir & "`", 3,  "batchLORASR")

; Call the main run function
LogMessage("Calling main batch function.", 2, "batchLORASR")
$iResult = BatchLORASR(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)

; Exit program
If ($iResult = 1) And (Not @error) Then
    LogMessage("Batch complete.", 1, "batchLORASR")
    Exit 0
Else
    LogMessage("Batch incomplete.", 1, "batchLORASR")
    Exit 1
EndIf
