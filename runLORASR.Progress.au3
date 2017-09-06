#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Progress
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.09.06
 Modified:       2017.09.06
 Version:        0.4.3.3

 Script Function:
    Show progress meters for the runLORASR batch process

#ce ----------------------------------------------------------------------------

#include-once
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <FontConstants.au3>
#include <WindowsConstants.au3>
#include <ProgressConstants.au3>
#include "runLORASR.Functions.au3"

; Code version
$g_sProgressVersion = "0.4.3.3"

; Handles of controls are global to allow updating at different stages of different processes
Global $g_hProgressWindow = 0, $g_hOverallProgressBar = 0, $g_hOverallProgressDetailLabel = 0, $g_hCurrentProgressBar = 0, $g_hCurrentProgressDetailLabel = 0

; Setting the value of $g_sProgressType determines whether calls will affect the overall or current progress bar
Global $g_sProgressType = "current"

Func DrawProgressWindow()
    LogMessage("Called `DrawProgressWindow()`", 5)

    ; Declarations
    Local $aiDesktopSize
    Local $iDesktopWidth, $iDesktopHeight, $iProgressWindowWidth, $iProgressWindowHeight, $iProgressWindowXPosition, $iProgressWindowYPosition
    Local $iPadding, $iLeft, $iTop, $iControlWidth, $iLabelHeight, $iProgressBarHeight
    Local $hProgressWindow = 0
    Local $hOverallProgressBar = 0, $hOverallProgressLabel = 0, $hOverallProgressDetailLabel = 0
    Local $hCurrentProgressBar = 0, $hCurrentProgressLabel = 0, $hCurrentProgressDetailLabel = 0

    ; Set window and control sizes and locations
    $iProgressWindowWidth = 300
    $iProgressWindowHeight = 165
    $iPadding = 5
    $iLeft = 2 * $iPadding
    $iTop = 2 * $iPadding
    $iControlWidth = $iProgressWindowWidth - (2 * $iLeft)
    $iLabelHeight = 15
    $iProgressBarHeight = 20

    ; Get desktop size
    LogMessage("Getting size of the desktop", 5, "DrawProgressWindow")
    $aiDesktopSize = WinGetPos("Program Manager")
    If UBound($aiDesktopSize) = 0 Or @error Then
        ThrowError("Error getting desktop size. Using default settings.", 5, "DrawProgressWindow", @error)
        $iProgressWindowXPosition = 20
        $iProgressWindowYPosition = 20
    Else
        $iDesktopWidth = $aiDesktopSize[2]
        $iDesktopHeight = $aiDesktopSize[3]
        $iProgressWindowXPosition = $iDesktopWidth - $iProgressWindowWidth - 50
        $iProgressWindowYPosition = $iDesktopHeight - $iProgressWindowHeight - 100
    EndIf

    ; Create progress window
    LogMessage("Creating the progress window", 5, "DrawProgressWindow")
    $hProgressWindow = GUICreate("runLORASR Progress", $iProgressWindowWidth, $iProgressWindowHeight, $iProgressWindowXPosition, $iProgressWindowYPosition, $WS_CAPTION & $WS_POPUP)
    If (Not $hProgressWindow) Or @error Then
        ThrowError("Error creating the progress window.", 3, "DrawProgressWindow", @error)
        Return 0
    EndIf

    ; Labels for overall progress bar
    LogMessage("Initializing the progress window", 5, "DrawProgressWindow")
    $hOverallProgressLabel = GUICtrlCreateLabel("Overall progress", $iLeft, $iTop, $iControlWidth, $iLabelHeight)
    GUICtrlSetFont($hOverallProgressLabel, 9, $FW_BOLD)
    $hOverallProgressDetailLabel = GUICtrlCreateLabel("Initializing...", $iLeft, $iTop + $iLabelHeight + $iPadding, $iControlWidth, $iLabelHeight)

    ; Overall progress bar
    $hOverallProgressBar = GUICtrlCreateProgress($iLeft, $iTop + (2 * $iLabelHeight) + (2 * $iPadding), $iControlWidth, $iProgressBarHeight)

    ; Labels for current action bar
    $hCurrentProgressLabel = GUICtrlCreateLabel("Current action", $iLeft, $iTop + (2 * $iLabelHeight) + (1 * $iProgressBarHeight) + (6 * $iPadding), $iControlWidth, $iLabelHeight)
    GUICtrlSetFont($hCurrentProgressLabel, 9, $FW_BOLD)
    $hCurrentProgressDetailLabel = GUICtrlCreateLabel("", $iLeft, $iTop + (3 * $iLabelHeight) + (1 * $iProgressBarHeight) + (7 * $iPadding), $iControlWidth, $iLabelHeight)

    ; Current action progress bar
    $hCurrentProgressBar = GUICtrlCreateProgress($iLeft, $iTop + (4 * $iLabelHeight) + (1 * $iProgressBarHeight) + (8 * $iPadding), $iControlWidth, $iProgressBarHeight, $PBS_SMOOTH)
    If @error Then
        ThrowError("Error initializing the progress window.", 3, "DrawProgressWindow", @error)
        Return 0
    EndIf

    ; Show window
    LogMessage("Displaying the progress window", 5, "DrawProgressWindow")
    GUISetState(@SW_SHOW, $hProgressWindow)
    If @error Then
        ThrowError("Error displaying the progress window.", 3, "DrawProgressWindow", @error)
        Return 0
    EndIf

    ; Save handles to global variables
    LogMessage("Saving progress window properties", 5, "DrawProgressWindow")
    $g_hProgressWindow = $hProgressWindow
    $g_hOverallProgressBar = $hOverallProgressBar
    $g_hOverallProgressDetailLabel = $hOverallProgressDetailLabel
    $g_hCurrentProgressBar = $hCurrentProgressBar
    $g_hCurrentProgressDetailLabel = $hCurrentProgressDetailLabel
    If @error Then
        ThrowError("Error saving progress window properties.", 3, "DrawProgressWindow", @error)
        Return 0
    EndIf

    ; Return window handle
    Return $hProgressWindow

EndFunc

Func CloseProgressWindow($hProgressWindow = $g_hProgressWindow)
    LogMessage("Called `CloseProgressWindow($hProgressWindow = " & hProgressWindow & ")`", 5)

    ; Close the window
    GUIDelete($hProgressWindow)
    If @error Then
        ThrowError("Error closing the progress window.", 4, "CloseProgressWindow", @error)
        Return 0
    EndIf

    ; Exit
    Return (Not @error)

EndFunc

Func UpdateProgress($sProgressType, $iProgressPercentage, $sLabel = "")
    LogMessage("Called `UpdateProgress($sProgressType = " & $sProgressType & ", $iProgressPercentage = " & $iProgressPercentage & ", $sLabel = " & $sLabel & ")`", 5)

    ; Declarations
    Local $hProgressBar = 0, $hProgressLabel = 0

    ; Get progress type
    Switch $sProgressType
        Case "overall"
            $hProgressBar = $g_hOverallProgressBar
            $hProgressLabel = $g_hOverallProgressDetailLabel
        Case "current"
            $hProgressBar = $g_hCurrentProgressBar
            $hProgressLabel = $g_hCurrentProgressDetailLabel
    EndSwitch
    If (Not $hProgressBar) Or @error Then
        ThrowError("Error selecting progress bar to update.", 5, "UpdateProgress", @error)
        Return 0
    EndIf

    ; Set progress
    LogMessage("Advancing " & $sProgressType & " progress bar to " & $iProgressPercentage & "%", 4, "UpdateProgress")
    GUICtrlSetData($hProgressBar, $iProgressPercentage)
    If @error Then
        ThrowError("Error updating progress bar.", 4, "UpdateProgress", @error)
    EndIf

    ; Set label
    If $hProgressLabel Then
        LogMessage("Progress update: " & $sLabel, 5, "UpdateProgress")
        GUICtrlSetData($hProgressLabel, $sLabel)
        If @error Then
            ThrowError("Error updating progress label.", 4, "UpdateProgress", @error)
        EndIf
    EndIf

    ; Exit
    Return (Not @error)

EndFunc

Func TestProgressBars()
    LogMessage("Called `TestProgressBars()`", 5)

    ; Declarations
    Local $i = 0, $j = 0, $iRuns = 0, $iTasks = 0

    ; How many runs and tasks to simulate
    $iRuns = 5
    $iTasks = 10
    LogMessage("Testing progress bars with " & $iRuns & " runs and " & $iTasks & " tasks.", 4, "UpdateProgress")

    ; Simple loop to check progress bars are both responding
    For $i = 1 To $iRuns
        Sleep(1000)
        UpdateProgress("overall", $i * 100/$iRuns, "Run " & $i & " of " & $iRuns)
        For $j = 1 To 10
            Sleep(200)
            UpdateProgress("current", $j * 100/$iTasks, "Task " & $j & " of " & $iTasks)
        Next
    Next

    ; Wait a bit
    Sleep(2000)

    ; Exit
    Return (Not @error)

EndFunc
