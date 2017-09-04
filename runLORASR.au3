#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Run LORASR for a given filename
#AutoIt3Wrapper_Res_Fileversion=0.4.2.3
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         runLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.04
 Modified:       2017.08.31
 Version:        0.4.2.3

Script Function:
    Run LORASR for a given filename

#ce ----------------------------------------------------------------------------

#include "runLORASR.Functions.au3"
#include "runLORASR.Run.au3"

LogMessage("Started `runLORASR` version 0.4.2.3", 3)

; Declarations
Local $iResult = 0
Local $sRun = "", $sNewFolder = "", $sSimulationProgramPath = ""

; Get global settings
LogMessage("Loading global settings...", 2, "runLORASR")
Local $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder
Local $bCleanup
$iResult = GetSettings(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)
If (Not $iResult) Or @error Then
    ThrowError("Error loading global settings", 1, "runLORASR", @error)
    Exit 2
EndIf

; Check command line parameters
LogMessage("Checking command line parameters...", 2, "runLORASR")
If $CmdLine[0] > 0 Then    $sRun = $CmdLine[1]
If @error Then
    ThrowError("Error checking command line parameters", 1, "runLORASR", @error)
    Exit 3
EndIf

; Get working directory if run from program directory
LogMessage("Checking current directory...", 2, "runLORASR")
If @WorkingDir = $sProgramPath Then
    $sNewFolder = FileSelectFolder("Please select the working directory in which to run LORASR.", "")
    If @error Then
        ThrowError("Could not select working directory", 1, "runLORASR", @error)
        Exit 4
    EndIf
    FileChangeDir($sNewFolder)
    If @error Then
        ThrowError("Could not access working directory", 1, "runLORASR", @error)
        Exit 5
    EndIf
EndIf

; Set up the simulation environment
LogMessage("Setting up simulation environment...", 2, "runLORASR")
$sSimulationProgramPath = SetupLORASR(@WorkingDir, $sProgramPath, $sSimulationProgram)
If (Not $sSimulationProgramPath) Or @error Then
    ThrowError("Could not set up simulation environment in folder `" & @WorkingDir & "`. Batch cancelled.", 1, "BatchLORASR", @error)
    Exit 6
EndIf

; Call the main run function
LogMessage("Starting run...", 1, "runLORASR")
$iResult = RunLORASR($sRun, @WorkingDir, $sSimulationProgramPath, $sInputFolder)
LogMessage("Run result: " & $iResult, 2, "runLORASR")

; Try once more if failed
If Not ($iResult = 1) Then
    LogMessage("Run failed. Trying again...", 1, "runLORASR")
    $iResult = RunLORASR($sRun, @WorkingDir, $sSimulationProgramPath, $sInputFolder)
    LogMessage("Run result: " & $iResult, 2, "runLORASR")
    If Not ($iResult = 1) Then
        LogMessage("Run failed again.", 1, "runLORASR")
        Exit 7
    EndIf
EndIf

; Exit program
LogMessage("Run complete.", 1, "runLORASR")
Exit
