#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Create a batch of input files from a batch definition file and a template
#AutoIt3Wrapper_Res_Fileversion=0.4.4.0
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         sweepLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.05
 Modified:       2017.09.08
 Version:        0.4.4.0

 Script Function:
    Create a batch of input files from a batch definition file and a template

#ce ----------------------------------------------------------------------------

; Load libraries
#include "runLORASR.Functions.au3"
#include "runLORASR.Sweep.au3"
#include "runLORASR.Progress.au3"

; Program version
Global CONST $g_sProgramName = "sweepLORASR"
Global CONST $g_sProgramVersion = "0.4.4.0"

; Declarations
Local $iResult = 0
Local $sFolder = ""

; Create log file
$g_sLogFile = DateTimeFileName($g_sProgramName, ".log.md")
CreateLogFile($g_sLogFile, @WorkingDir)
LogVersions($g_sProgramName, $g_sProgramVersion)

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

; Draw and show the progress meters
DrawProgressWindow()
$g_sProgressType = "both"

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
