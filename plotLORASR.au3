#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Copy data from LORASR output files to Excel plotting spreadsheet
#AutoIt3Wrapper_Res_Fileversion=0.4.4.0
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         plotLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.04
 Modified:       2017.09.08
 Version:        0.4.4.0

 Script Function:
    Copy data from LORASR output files to Excel plotting spreadsheet

#ce ----------------------------------------------------------------------------

; Load libraries
#include "runLORASR.Functions.au3"
#include "runLORASR.Plots.au3"

; Program version
Global CONST $g_sProgramName = "plotLORASR"
Global CONST $g_sProgramVersion = "0.4.4.0"

; Declarations
Local $iResult = 0
Local $sMasterPlotFile = "", $sPlotFile = "", $sFolder = ""

; Create log file
$g_sLogFile = DateTimeFileName($g_sProgramName, ".log.md")
CreateLogFile($g_sLogFile, @WorkingDir)
LogVersions($g_sProgramName, $g_sProgramVersion)

; Get global settings
LogMessage("Loading global settings...", 2, "plotLORASR")
Local $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder
Local $bCleanup
$iResult = GetSettings(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sMasterPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)
If (Not $iResult) Or @error Then
    ThrowError("Error loading global settings", 1, "plotLORASR", @error)
    Exit 1
EndIf

; Check command line parameters
LogMessage("Checking command line parameters...", 2, "plotLORASR")
If $CmdLine[0] > 0 Then
    $sPlotFile = $CmdLine[1] & ".xlsx"
Else
    $sPlotFile = $sMasterPlotFile
EndIf
If @error Then
    ThrowError("Error checking command line parameters", 1, "plotLORASR", @error)
    Exit 2
EndIf

; Get working directory if run from program directory
LogMessage("Checking current directory...", 2, "plotLORASR")
If @WorkingDir = $sProgramPath Then
    $sFolder = FileSelectFolder("Please select the working directory that holds the files to be plotted.", "")
    If @error Then
        ThrowError("Could not select working directory", 1, "plotLORASR", @error)
        Exit 3
    EndIf
    FileChangeDir($sFolder)
    If @error Then
        ThrowError("Could not access working directory", 1, "plotLORASR", @error)
        Exit 4
    EndIf
EndIf

; Call the main run function
LogMessage("Starting plots...", 1, "plotLORASR")
$iResult = PlotLORASR(@WorkingDir, $sPlotFile, $sMasterPlotFile, $sProgramPath)
If (Not $iResult) Or @error Then
    ThrowError("Error during plot process.", 1, "plotLORASR", @error)
    Exit 5
EndIf

; Exit program
LogMessage("Plots complete.", 1, "plotLORASR")
Exit
