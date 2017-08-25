#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Comment=Created by M J Easton
#AutoIt3Wrapper_Res_Description=Tidy up files from a batch run of LORASR
#AutoIt3Wrapper_Res_Fileversion=0.4.1.0
#AutoIt3Wrapper_Res_LegalCopyright=Creative Commons Attribution ShareAlike
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 Script:         tidyLORASR
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.13
 Modified:       2017.08.11
 Version:        0.4.1.0

 Script Function:
	Tidy up files from a batch run of LORASR

#ce ----------------------------------------------------------------------------

#include "runLORASR.Functions.au3"
#include "runLORASR.Tidy.au3"

LogMessage("Started tidyLORASR version 0.4.1.0", 2)

; Declarations
Local $sRun = ""
Local $bError = False

; Get global settings
LogMessage("Loading global settings...", 2, "tidyLORASR")
Local $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder
Local $bCleanup
If Not GetSettings(@WorkingDir, $sProgramPath, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup) Then
	ThrowError("Could not get global settings. Exiting program.", 1, "tidyLORASR", @error)
	Exit 2
EndIf

; Don't do anything if not meant to clean up
If Not $bCleanup Then
	LogMessage("Cleanup option switch off, exiting.", 2, "tidyLORASR")
	Exit 0
EndIf

; Get working directory if run from program directory
LogMessage("Checking current directory...", 2, "tidyLORASR")
If @WorkingDir = $sProgramPath Then
	$sRun = FileSelectFolder("Please select the working directory that holds the files to be cleaned up.", "")
	If @error Then
		ThrowError("Could not select working directory", 1, "tidyLORASR", @error)
		Exit 3
	EndIf
	FileChangeDir($sRun)
	If @error Then
		ThrowError("Could not access working directory", 1, "tidyLORASR", @error)
		Exit 4
	EndIf
EndIf

; Tidy up input files for each run: completed runs go back in the input folder, incomplete runs are moved to an incomplete folder for later processing.
LogMessage("Sorting through any input files...", 2, "tidyLORASR")
If Not TidyAllRunFiles(@WorkingDir, $sInputFolder, $sIncompleteFolder) Then $bError = True

; Tidy up other files
LogMessage("Clearing up any other files...", 2, "tidyLORASR")
If Not TidyBatchFiles(@WorkingDir, $sSimulationProgram, $sSweepFile, $sTemplateFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder) Then $bError = True

; Report completion
If @error or $bError Then
	LogMessage("Finished tidying with some errors.", 1, "tidyLORASR")
Else
	LogMessage("Finished tidying.", 1, "tidyLORASR")
EndIf

; End program
Exit (Not (@error or $bError))