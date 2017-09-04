#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Create a batch of input files from a batch definition file and a template
#AutoIt3Wrapper_Res_Fileversion=0.4.2.3
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         sweepLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.05
 Modified:       2017.08.31
 Version:        0.4.2.3

 Script Function:
    Create a batch of input files from a batch definition file and a template

#ce ----------------------------------------------------------------------------

#include "runLORASR.Functions.au3"
#include "runLORASR.Sweep.au3"

LogMessage("Started `sweepLORASR` version 0.4.2.3", 3)

; Declarations
Local $iResult = 0
Local $sFolder = ""

; Get global settings
LogMessage("Loading global settings...", 2, "sweepLORASR")
Local $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder
Local $bCleanup
$iResult = GetSettings(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)
If (Not $iResult) Or @error Then
    ThrowError("Error loading global settings", 1, "sweepLORASR", @error)
    Exit 1
EndIf

; Get working directory if run from program directory
LogMessage("Checking current directory...", 2, "sweepLORASR")
If @WorkingDir = $sProgramPath Then
    $sFolder = FileSelectFolder("Please select the working directory that holds the sweep definition files.", "")
    If @error Then
        ThrowError("Could not select working directory", 1, "sweepLORASR", @error)
        Exit 2
    EndIf
    FileChangeDir($sFolder)
    If @error Then
        ThrowError("Could not access working directory", 1, "sweepLORASR", @error)
        Exit 3
    EndIf
EndIf

; Copy sweep files if in subfolder
LogMessage("Checking subfolders...", 2, "sweepLORASR")

; Call the main run function
LogMessage("Starting preparations for parameter sweep...", 2, "sweepLORASR")
$iResult = SweepLORASR(@WorkingDir, $sSweepFile, $sTemplateFile, $sResultsFile, $sInputFolder)
If (Not $iResult) Or @error Then
    ThrowError("Error during parameter sweep.", 1, "sweepLORASR", @error)
    Exit 4
EndIf

; Exit program
LogMessage("Sweep complete.", 1, "sweepLORASR")
Exit
