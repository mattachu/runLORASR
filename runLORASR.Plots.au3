#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Plots
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.13
 Modified:       2017.08.09
 Version:        0.4.0.72

 Script Function:
	Copy data from LORASR output files to Excel plotting spreadsheet

#ce ----------------------------------------------------------------------------

#include-once
#include "runLORASR.Functions.au3"
#include <Excel.au3>
#include <Date.au3>

LogMessage("Loaded runLORASR.Plots version 0.4.0.72", 3)

; Function that defines input and output settings for each data type: changes to the LORASR code may require changes here.
Func GetFileSettings($sDataType, ByRef $sDataFile1, ByRef $iDataStart1, ByRef $iDataEnd1, ByRef $sDataFile2, ByRef $iDataStart2, ByRef $iDataEnd2, ByRef $iWorksheet, ByRef $sDataLocation1, ByRef $sDataLocation2)
	LogMessage("Called GetFileSettings($sDataType = " & $sDataType & ", ByRef $sDataFile1, ByRef $iDataStart1, ByRef $iDataEnd1, ByRef $sDataFile2, ByRef $iDataStart2, ByRef $iDataEnd2, ByRef $iWorksheet, ByRef $sDataLocation1, ByRef $sDataLocation2)", 5)

	Switch $sDataType
		Case "emittance values"
			$sDataFile1 = "emival"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iWorksheet = 12
			$sDataLocation1 = "D2"
		Case "transverse envelopes"
			$sDataFile1 = "transenv"
			$iDataStart1 = 1
			$iDataEnd1 = 5
			$iWorksheet = 1
			$sDataLocation1 = "B2"
		Case "longitudinal envelopes"
			$sDataFile1 = "longen"
			$sDataFile2 = "longph"
			$iDataStart1 = 1
			$iDataEnd1 = 3
			$iDataStart2 = 2
			$iDataEnd2 = 3
			$iWorksheet = 2
			$sDataLocation1 = "B2"
			$sDataLocation2 = "E2"
		Case "input x emittance"
			$sDataFile1 = "xcluin"
			$sDataFile2 = "xclueli"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iDataStart2 = 1
			$iDataEnd2 = 3
			$iWorksheet = 3
			$sDataLocation1 = "A2"
			$sDataLocation2 = "D2"
		Case "output x emittance"
			$sDataFile1 = "xcluout"
			$sDataFile2 = "xcluelo"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iDataStart2 = 1
			$iDataEnd2 = 3
			$iWorksheet = 4
			$sDataLocation1 = "A2"
			$sDataLocation2 = "D2"
		Case "input y emittance"
			$sDataFile1 = "ycluin"
			$sDataFile2 = "yclueli"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iDataStart2 = 1
			$iDataEnd2 = 3
			$iWorksheet = 5
			$sDataLocation1 = "A2"
			$sDataLocation2 = "D2"
		Case "output y emittance"
			$sDataFile1 = "ycluout"
			$sDataFile2 = "ycluelo"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iDataStart2 = 1
			$iDataEnd2 = 3
			$iWorksheet = 6
			$sDataLocation1 = "A2"
			$sDataLocation2 = "D2"
		Case "input z emittance"
			$sDataFile1 = "zcluin"
			$sDataFile2 = "zclueli"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iDataStart2 = 1
			$iDataEnd2 = 3
			$iWorksheet = 7
			$sDataLocation1 = "A2"
			$sDataLocation2 = "D2"
		Case "output z emittance"
			$sDataFile1 = "zcluout"
			$sDataFile2 = "zcluelo"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iDataStart2 = 1
			$iDataEnd2 = 3
			$iWorksheet = 8
			$sDataLocation1 = "A2"
			$sDataLocation2 = "D2"
		Case "bunch center"
			$sDataFile1 = "bucent"
			$iDataStart1 = 1
			$iDataEnd1 = 2
			$iWorksheet = 9
			$sDataLocation1 = "A2"
		Case "emittance growth"
			$sDataFile1 = "rmsgrow"
			$iDataStart1 = 1
			$iDataEnd1 = 4
			$iWorksheet = 10
			$sDataLocation1 = "B2"
		Case "phase advance"
			$sDataFile1 = "phadv"
			$iDataStart1 = 1
			$iDataEnd1 = 4
			$iWorksheet = 11
			$sDataLocation1 = "A2"
		Case Else
			ThrowError("Plot data type '" & $sDataType & "' not found.", "GetFileSettings", @error)
			SetError(1)
			Return 0
	EndSwitch

	; Exit
	Return (Not @error)

EndFunc

; Main function
Func PlotLORASR($sWorkingDirectory = @WorkingDir, $sPlotFile = "Plots.xlsx", $sMasterPlotFile = "Plots.xlsx", $sMasterPath = "C:\Program Files (x86)\LORASR")
	LogMessage("Called PlotLORASR($sWorkingDirectory = " & $sWorkingDirectory & ", $sPlotFile = " & $sPlotFile & ", $sMasterPlotFile = " & $sMasterPlotFile & ", $sMasterPath = " & $sMasterPath & ")", 5)

	; Declarations
	Local $oExcel, $oWorkbook
	Local $sSpreadsheet = ""
	Local $iResult = 0

	; Create the new workbook
	LogMessage("Creating new workbook", 3, "PlotLORASR")
	$sSpreadsheet = CreatePlotSpreadsheet($sWorkingDirectory, $sPlotFile, $sMasterPlotFile, $sMasterPath)
	If (Not $sSpreadsheet) Or @error Then
		ThrowError("Could not create the plot spreadsheet, cannot continue with plotting.", 2, "PlotLORASR", @error)
		SetError(1)
		Return 0
	Else
		LogMessage("Workbook created at " & $sSpreadsheet, 3, "PlotLORASR")
	EndIf

	; Start Excel and open the workbook
	LogMessage("Opening workbook", 3, "PlotLORASR")
	$iResult = OpenPlotSpreadsheet($oExcel, $oWorkbook, $sSpreadsheet)
	If (Not $iResult) Or @error Then
		ThrowError("Could not open the plot spreadsheet, cannot continue with plotting.", 2, "PlotLORASR", @error)
		SetError(2)
		Return 0
	EndIf

	; Plot the data
	LogMessage("Plotting data into workbook...", 3, "PlotLORASR")
	$iResult = PlotAllData($oWorkbook, $sWorkingDirectory)
	If (Not $iResult) Or @error Then
		ThrowError("Error " & @error & " while plotting, attempting to save and close plot spreadsheet.", 3, "PlotLORASR", @error)
		SetError(3)
	EndIf

	; Save and close the spreadsheet
	LogMessage("Closing workbook...", 3, "PlotLORASR")
	$iResult = SaveAndClosePlotSpreadsheet($oExcel, $oWorkbook)
	If (Not $iResult) Or @error Then
		ThrowError("Could not save and close Excel workbook.", 2, "PlotLORASR", @error)
		SetError(4)
		Return 0
	EndIf

	; Exit
	LogMessage("Plot process concluded.", 3, "PlotLORASR")
	Return (Not @error)

EndFunc

; Function to plot data for each output file from LORASR
Func PlotAllData(ByRef $oWorkbook, $sWorkingDirectory = @WorkingDir)
	LogMessage("Called PlotAllData(ByRef $oWorkbook, $sWorkingDirectory = " & $sWorkingDirectory & ")", 5)

	; Emittance values
	LogMessage("Plotting emittance values...", 4, "PlotAllData")
	If Not PlotData("emittance values", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error setting emittance values. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(1)
	EndIf

	; Transverse envelope data
	LogMessage("Plotting transverse envelope data...", 4, "PlotAllData")
	If Not PlotData("transverse envelopes", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting transverse envelopes. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(2)
	EndIf

	; Longitudinal envelope data
	LogMessage("Plotting longitudinal envelope data...", 4, "PlotAllData")
	If Not PlotData("longitudinal envelopes", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting longitudinal envelopes. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(3)
	EndIf

	; Input x-emittance data
	LogMessage("Plotting input x emittance data...", 4, "PlotAllData")
	If Not PlotData("input x emittance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting input x emittance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(4)
	EndIf

	; Output x-emittance data
	LogMessage("Plotting output x emittance data...", 4, "PlotAllData")
	If Not PlotData("output x emittance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting output x emittance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(5)
	EndIf

	; Input y-emittance data
	LogMessage("Plotting input y emittance data...", 4, "PlotAllData")
	If Not PlotData("input y emittance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting input y emittance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(6)
	EndIf

	; Output y-emittance data
	LogMessage("Plotting output y emittance data...", 4, "PlotAllData")
	If Not PlotData("output y emittance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting output y emittance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(7)
	EndIf

	; Input z-emittance data
	LogMessage("Plotting input z emittance data...", 4, "PlotAllData")
	If Not PlotData("input z emittance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting input z emittance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(8)
	EndIf

	; Output z-emittance data
	LogMessage("Plotting output z emittance data...", 4, "PlotAllData")
	If Not PlotData("output z emittance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting output z emittance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(9)
	EndIf

	; Bunch center data
	LogMessage("Plotting bunch center data...", 4, "PlotAllData")
	If Not PlotData("bunch center", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting bunch center data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(10)
	EndIf

	; Emittance growth data
	LogMessage("Plotting emittance growth data...", 4, "PlotAllData")
	If Not PlotData("emittance growth", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting emittance growth data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(11)
	EndIf

	; Phase advance data
	LogMessage("Plotting phase advance data...", 4, "PlotAllData")
	If Not PlotData("phase advance", $oWorkbook, $sWorkingDirectory) Then
		ThrowError("Error plotting phase advance data. Continuing with plots.", 3, "PlotAllData", @error)
		SetError(12)
	EndIf

	; Exit
	If @error Then
		LogMessage("Plots completed, but with errors.", 3, "PlotAllData")
		Return 0
	Else
		LogMessage("Plots completed.", 3, "PlotAllData")
		Return 1
	EndIf

EndFunc

; Function to plot a given set of data into the give Excel workbook
Func PlotData($sDataType, ByRef $oWorkbook, $sWorkingDirectory = @WorkingDir)
	LogMessage("Called PlotData($sDataType = " & $sDataType & ", ByRef $oWorkbook, $sWorkingDirectory = " & $sWorkingDirectory & ")", 5)

	; Declarations
	Local $sDataFile1 = "", $sDataFile2 = "", $sDataLocation1 = "", $sDataLocation2 = ""
	Local $iWorksheet = 0, $iDataStart1 = 0, $iDataStart2 = 0, $iDataEnd1 = 0, $iDataEnd2 = 0, $iResult = 0
	Local $asData

	; Get file settings
	LogMessage("Getting file settings for " & $sDataType, 5, "PlotData")
	$iResult = GetFileSettings($sDataType, $sDataFile1, $iDataStart1, $iDataEnd1, $sDataFile2, $iDataStart2, $iDataEnd2, $iWorksheet, $sDataLocation1, $sDataLocation2)
	If (Not $iResult) Or @error Then
		ThrowError("Could not set plot data type.", 3, "PlotData", @error)
		SetError(1)
		Return 0
	EndIf

	; Check all parameters are set for file 1
	If Not $sDataFile1 Then
		ThrowError("Plot data file not defined.", 3, "PlotData", @error)
		SetError(2)
		Return 0
	EndIf
	If Not ($iWorksheet And $sDataLocation1) Then
		ThrowError("Data write location not defined.", 3, "PlotData", @error)
		SetError(3)
		Return 0
	EndIf

	; Load data from file 1
	LogMessage("Loading data from " & $sDataFile1, 5, "PlotData")
	$asData = LoadData($sDataFile1, $sWorkingDirectory, $iDataStart1, $iDataEnd1)
	If @error Then
		ThrowError("Error while loading data in file " & $sDataFile1, 3, "PlotData", @error)
		SetError(4)
		Return 0
	EndIf

	; Write data from file 1 to spreadsheet
	LogMessage("Writing data from " & $sDataFile1 & " to workbook", 5, "PlotData")
	WriteData($oWorkbook, $iWorksheet, $sDataLocation1, $asData)
	If @error Then
		ThrowError("Error while writing data in file " & $sDataFile1, 3, "PlotData", @error)
		SetError(5)
		Return 0
	EndIf

	; Process file 2 if required
	If $sDataFile2 And $iDataStart2 And $iDataEnd2 And $iWorksheet And $sDataLocation2 Then

		; Load data from file 2
		LogMessage("Loading data from " & $sDataFile2, 5, "PlotData")
		$asData = LoadData($sDataFile2, $sWorkingDirectory, $iDataStart2, $iDataEnd2)
		If @error Then
			ThrowError("Error while loading data in file " & $sDataFile2, 3, "PlotData", @error)
			SetError(6)
			Return 0
		EndIf

		; Write data from file 2 to spreadsheet
		LogMessage("Writing data from " & $sDataFile2 & " to workbook", 5, "PlotData")
		WriteData($oWorkbook, $iWorksheet, $sDataLocation2, $asData)
		If @error Then
			ThrowError("Error while writing data in file " & $sDataFile2, 3, "PlotData", @error)
			SetError(7)
			Return 0
		EndIf

	EndIf

	; Exit
	Return (Not @error)

EndFunc

; Function to create a new spreadsheet from the master; returns the full path and filename as a string
Func CreatePlotSpreadsheet($sWorkingDirectory = @WorkingDir, $sPlotFile = "Plots.xlsx", $sMasterPlotFile = "Plots.xlsx", $sMasterPath = "C:\Program Files (x86)\LORASR")
	LogMessage("Called CreatePlotSpreadsheet($sWorkingDirectory = " & $sWorkingDirectory & ", $sPlotFile = " & $sPlotFile & ", $sMasterPlotFile = " & $sMasterPlotFile & ", $sMasterPath = " & $sMasterPath & ")", 5)

	; Declarations
	Local $tCurrentTime

	; If there is no special name for the plot file, use the date and time
	If $sPlotFile = $sMasterPlotFile Then
		$tCurrentTime = _Date_Time_GetLocalTime()
		$sPlotFile = StringTrimRight($sPlotFile, 5) & "-" & StringReplace(StringReplace(StringReplace(_Date_Time_SystemTimeToDateTimeStr($tCurrentTime,1), "/", ""), ":", ""), " ", "-") & ".xlsx"
		LogMessage("Creating unique plot file name: " & $sPlotFile, 5, "CreatePlotSpreadsheet")
	EndIf

	; Find master spreadsheet
	LogMessage("Searching for master plot spreadsheet", 5, "CreatePlotSpreadsheet")
	If FileExists($sWorkingDirectory & "\" & $sMasterPlotFile) Then
		; Found in working directory
		$sMasterPlotFile = $sWorkingDirectory & "\" & $sMasterPlotFile
	ElseIf FileExists($sMasterPath & "\" & $sMasterPlotFile) Then
		; Found in program directory
		$sMasterPlotFile = $sMasterPath & "\" & $sMasterPlotFile
	Else
		; Not found
		ThrowError("Master plot spreadsheet not found.", "plotLORASR", @error)
		SetError(1)
		Return 0
	EndIf
	If @error Then
		; Error while searching
		ThrowError("Could not locate master plot spreadsheet.", 3, "plotLORASR", @error)
		SetError(2)
		Return 0
	EndIf

	; Define plot file location
	$sPlotFile = $sWorkingDirectory & "\" & $sPlotFile

	; Move existing files out of the way
	If FileExists($sPlotFile) Then FileMove($sPlotFile, $sPlotFile & ".old", $FC_OVERWRITE)
	If @error Then
		ThrowError("Could not create plot spreadsheet.", 3, "plotLORASR", @error)
		SetError(3)
		Return 0
	EndIf

	; Create plot spreadsheet from master
	If Not FileCopy($sMasterPlotFile, $sPlotFile) Then
		ThrowError("Could not create plot spreadsheet.", 3, "plotLORASR", @error)
		SetError(4)
		Return 0
	EndIf

	; Exit
	Return $sPlotFile

EndFunc

; Function to open a given spreadsheet and return handles for the Excel object and workbook object
Func OpenPlotSpreadsheet(ByRef $oExcel, ByRef $oWorkbook, $sSpreadsheet = @WorkingDir & "\Plots.xlsx")
	LogMessage("Called OpenPlotSpreadsheet(ByRef $oExcel, ByRef $oWorkbook, $sSpreadsheet = " & $sSpreadsheet & ")", 5)

	; Open Excel
	$oExcel = _Excel_Open(False)
	If @error Then
		ThrowError("Could not create the Excel application object.", 3, "OpenPlotSpreadsheet", @error)
		SetError(1)
		Return 0
	EndIf

	; Open the spreadsheet
	$oWorkbook = _Excel_BookOpen($oExcel, $sSpreadsheet)
	If @error Then
		ThrowError("Could not open Excel spreadsheet.", 3, "OpenPlotSpreadsheet", @error)
		SetError(1)
		Return 0
	EndIf

	; Exit
	Return (Not @error)

EndFunc

; Function to save workbook and quit Excel based on object handles
Func SaveAndClosePlotSpreadsheet(ByRef $oExcel, ByRef $oWorkbook)
	LogMessage("Called SaveAndClosePlotSpreadsheet(ByRef $oExcel, ByRef $oWorkbook)", 5)

	; Save the spreadsheet
	_Excel_BookSave($oWorkbook)
	If @error Then
		ThrowError("Could not save workbook.", 3, "SaveAndClosePlotSpreadsheet", @error)
		SetError(1)
		Return 0
	EndIf

	; Wait for save to complete
	Sleep(50)

	; Close Excel
	_Excel_Close($oExcel)
	If @error Then
		ThrowError("Could not close Excel.", 3, "SaveAndClosePlotSpreadsheet", @error)
		SetError(2)
		Return 0
	EndIf

	; Wait for Excel to exit
	Sleep(50)

	; Exit
	Return (Not @error)

EndFunc

; Function to load data from file
Func LoadData($sDataFile, $sWorkingDirectory, $iDataStart, $iDataEnd)
	LogMessage("Called LoadData($sDataFile = " & $sDataFile & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $iDataStart = " & $iDataStart & ", $iDataEnd = " & $iDataEnd & ")", 5)

	; Declarations
	Local $iDataCount = 0, $iCurrentLine = 0, $iCurrentDataPoint = 0

	; Load data from file
	Local $asFileData = FileReadToArray($sWorkingDirectory & "\" & $sDataFile)
	If @error Then
		ThrowError("Error while loading data from file " & $sDataFile, 3, "LoadData", @error)
		SetError(1)
		Return 0
	EndIf

	; Number of data columns
	$iDataCount = $iDataEnd - $iDataStart + 1

	; Build data array from file data
	Local $asData[UBound($asFileData)][$iDataCount]
	For $iCurrentLine = 0 to UBound($asFileData) - 1
		$sFileLine = $asFileData[$iCurrentLine]
		; Remove leading and trailing whitespace
		$sFileLine = StringStripWS($sFileLine, $STR_STRIPLEADING)
		$sFileLine = StringStripWS($sFileLine, $STR_STRIPTRAILING)
		; Replace internal whitespace with commas
		$sFileLine = StringRegExpReplace($sFileLine, "\h+", ",")
		; Split line into array
		$sFileLine = StringSplit($sFileLine, ",")
		; Save to data array
		For $iCurrentDataPoint = $iDataStart To $iDataEnd
			$asData[$iCurrentLine][$iCurrentDataPoint-$iDataStart] = $sFileLine[$iCurrentDataPoint]
		Next ;$iCurrentDataPoint
	Next ;$iCurrentLine
	If @error Then
		ThrowError("Error while processing data in file " & $sDataFile, 3, "LoadData", @error)
		SetError(2)
		Return 0
	EndIf

	Return $asData

EndFunc

; Function to write data to spreadsheet
Func WriteData(ByRef $oWorkbook, $iWorksheet, $sDataLocation, $asData)
	LogMessage("Called WriteData(ByRef $oWorkbook, $iWorksheet = " & $iWorksheet & ", $sDataLocation = " & $sDataLocation & ", $asData = <array>)", 5)

	; Write given data to given location in workbook
	_Excel_RangeWrite($oWorkbook, $iWorksheet, $asData, $sDataLocation)
	If @error Then
		ThrowError("Error writing to spreadsheet.", 3, "WriteData", @error)
		SetError(1)
		Return 0
	EndIf

	Return (Not @error)

EndFunc