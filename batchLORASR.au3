#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Work through a batch of input files and run LORASR for each one
#AutoIt3Wrapper_Res_Fileversion=0.4.2.4
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         batchLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.04
 Modified:       2017.09.04
 Version:        0.4.2.4

 Script Function:
    Work through a batch of input files and run LORASR for each one

#ce ----------------------------------------------------------------------------

#include "runLORASR.Functions.au3"
#include "runLORASR.Batch.au3"

LogMessage("Starting `batchLORASR` version 0.4.2.4", 3)

; Declarations
Local $iResult = 0
Local $sNewFolder = ""

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
LogMessage("Running batch in folder `" & @WorkingDir & "`", 3,  "batchLORASR")

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
