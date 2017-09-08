#cs ----------------------------------------------------------------------------

 Script:         runLORASR.Run
 AutoIt Version: 3.3.14.2
 Author:         Matt Easton
 Created:        2017.07.04
 Modified:       2017.09.08
 Version:        0.4.4.0

 Script Function:
    Run LORASR for a given filename

#ce ----------------------------------------------------------------------------

#include-once
#include "runLORASR.Functions.au3"

; Code version
$g_sRunVersion = "0.4.4.0"

; Global declarations
Global $g_sMainWindowTitle = "LORASR PC Version"
Global $g_sLoadInputFileWindowTitle = "Load Input Data File"
Global $g_sConsoleWindowTitle = "Console Output"
Global $g_sGenericWindowTitle = "LORASR"

; Preparation
Func SetupLORASR($sWorkingDirectory = @WorkingDir, $sProgramPath = "C:\Program Files (x86)\LORASR", $sSimulationProgram = "LORASR.exe")
    LogMessage("Called `SetupLORASR($sWorkingDirectory = " & $sWorkingDirectory & ", $sProgramPath = " & $sProgramPath & ", $sSimulationProgram = " & $sSimulationProgram & ")`", 5)

    ; Declarations
    Local $sSimulationProgramPath = ""

    ; Find executable
    LogMessage("Searching for LORASR executable", 3, "SetupLORASR")
    $sSimulationProgramPath = FindFile($sSimulationProgram, $sWorkingDirectory, $sProgramPath, True)
    If Not $sSimulationProgramPath Then
        ThrowError("LORASR executable not found", 2, "SetupLORASR", @error)
        SetError(1)
        Return 0
    EndIf

    ; Return the location of the program
    Return $sSimulationProgramPath

EndFunc

; Main function
Func RunLORASR($sRun = "", $sWorkingDirectory = @WorkingDir, $sSimulationProgramPath = $sWorkingDirectory & "\LORASR.exe", $sInputFolder = "Input")
    LogMessage("Called `RunLORASR($sRun = " & $sRun & ", $sWorkingDirectory = " & $sWorkingDirectory & ", $sSimulationProgramPath = " & $sSimulationProgramPath & ", $sInputFolder = " & $sInputFolder & ")`", 5)

    ; Declarations
    Local $sInputFile = ""
    Local $hLORASR = 0

    ; Get input file name
    LogMessage("Selecting input file...", 3, "RunLORASR")
    If Not $sRun Then
        $sInputFile = InputBox("runLORASR", "Enter the filename of the LORASR input file:")
        If @error Then
            ThrowError("No input file found.", 2, "RunLORASR", @error)
            SetError(2, @error)
            Return 0
        EndIf
    Else
        $sInputFile = $sRun & ".in"
    EndIf

    ; Find input file
    $sInputFile = FindFile($sInputFile, $sWorkingDirectory, $sWorkingDirectory & "\" & $sInputFolder, True)
    If Not $sInputFile Then
        ThrowError("Input file not found.", 2, "RunLORASR", @error)
        SetError(3)
        Return 0
    EndIf

    ; Close any existing LORASR process
    LogMessage("Closing existing LORASR processes...", 3, "RunLORASR")
    If WinExists($g_sGenericWindowTitle) Then KillLORASR()
    If @error Then
        ThrowError("Could not close existing LORASR processes. Attempting to continue...", 3, "RunLORASR", @error)
        SetError(0)
    EndIf

    ; Start the LORASR software
    LogMessage("Launching LORASR executable...", 3, "RunLORASR")
    Run($sSimulationProgramPath)
    If @error Then
        ThrowError("Could not run LORASR.", 2, "RunLORASR", @error)
        KillLORASR()
        SetError(4)
        Return 0
    EndIf

    ; Open the input file
    LogMessage("Loading input file...", 3, "RunLORASR")
    $hLORASR = LoadInputFile($sInputFile)
    If $hLORASR = 0 Then
        ThrowError("Could not load the input file.", 2, "RunLORASR", @error)
        KillLORASR()
        SetError(5)
        Return 0
    EndIf

    ; Run the calculation
    LogMessage("Running calculation...", 3, "RunLORASR")
    $hLORASR = RunCalculation()
    If $hLORASR = 0 Then
        ThrowError("Could not run the simulation.", 2, "RunLORASR", @error)
        KillLORASR()
        SetError(6)
        Return 0
    EndIf

    ; Quit the program
    LogMessage("Closing LORASR...", 3, "RunLORASR")
    If Not WinClose($g_sMainWindowTitle) Then
        ThrowError("Could not close LORASR.", 3, "RunLORASR", @error)
        ; Try to force LORASR to close
        If Not KillLORASR() Then
            SetError(7)
            Return 0
        EndIf
    EndIf

    ; Quit the program
    LogMessage("Simulation complete: exiting `RunLORASR`", 2, "RunLORASR")
    Return 1

EndFunc   ;==>RunLORASR

; Function to load an input file
Func LoadInputFile($sInputFile)
    LogMessage("Called `LoadInputFile($sInputFile = " & $sInputFile & ")`", 5)

    ; Declarations
    Local $hLORASR = 0

    ; Activate the main window
    LogMessage("Activating LORASR main window.", 5, "LoadInputFile")
    $hLORASR = SafeActivate($g_sMainWindowTitle)
    If $hLORASR = 0 Then
        ThrowError("Could not activate LORASR.", 3, "LoadInputFile", @error)
        SetError(1)
        Return 0
    EndIf

    ; Send keystrokes to File > Open
    LogMessage("Sending open file signal.", 5, "LoadInputFile")
    Send("!f")
    Send("o")

    ; Wait for file open dialog to appear
    LogMessage("Activating file open dialog window.", 5, "LoadInputFile")
    $hLORASR = SafeActivate($g_sLoadInputFileWindowTitle)
    If $hLORASR = 0 Then
        ThrowError("File open dialog not responding.", 3, "LoadInputFile", @error)
        SetError(2)
        Return 0
    EndIf

    ; Fill in the full path and filename
    Sleep(100)
    LogMessage("Sending input file path.", 5, "LoadInputFile")
    Send($sInputFile)
    Send("{ENTER}")

    ; Check for errors
    LogMessage("Checking for success.", 5, "LoadInputFile")
    If WinActivate($g_sLoadInputFileWindowTitle) Then
        ; Try to cancel process and retry
        LogMessage("Could not load the input file. Trying again...", 3, "LoadInputFile")
        If CancelLoadInputFile() Then
            ; Try again
            LogMessage("Activating LORASR main window.", 5, "LoadInputFile")
            SafeActivate($g_sMainWindowTitle)
            LogMessage("Sending open file signal.", 5, "LoadInputFile")
            Send("!f")
            Send("o")
            LogMessage("Activating file open dialog window.", 5, "LoadInputFile")
            SafeActivate($g_sLoadInputFileWindowTitle)
            Sleep(100)
            LogMessage("Sending input file path.", 5, "LoadInputFile")
            Send($sInputFile)
            Send("{ENTER}")
            ; Check for success
            LogMessage("Checking for success.", 5, "LoadInputFile")
            If Not WinActive($g_sMainWindowTitle) Then
                ThrowError("Could not load the input file after multiple attempts.", 3, "LoadInputFile", @error)
                SetError(4)
                Return 0
            EndIf
        Else
            ThrowError("Error while cancelling load.", 3, "LoadInputFile", @error)
            SetError(3)
            Return 0
        EndIf
    EndIf

    ; Exit
    LogMessage("Input file loaded.", 4, "LoadInputFile")
    Return $hLORASR

EndFunc

; Function to run the calculation in LORASR
Func RunCalculation()
    LogMessage("Called RunCalculation()", 5)

    ; Declarations
    Local $hLORASR = 0
    Local $iCalculationTimeout = 60

    ; Activate the main window
    LogMessage("Activating LORASR main window.", 5, "RunCalculation")
    $hLORASR = SafeActivate($g_sMainWindowTitle)
    If $hLORASR = 0 Then
        ThrowError("Could not activate LORASR.", 3, "RunCalculation", @error)
        SetError(1)
        Return 0
    EndIf

    ; Send keystrokes to Run > Calculation
    LogMessage("Sending start calculation signal.", 5, "RunCalculation")
    Send("!r")
    Send("c")

    ; Wait for the calculation window to be complete (transmission ratio is the last line of ouptut)
    LogMessage("Activating LORASR console window and waiting for run result.", 5, "RunCalculation")
    $hLORASR = SafeActivate($g_sConsoleWindowTitle, "TRATIO=", $iCalculationTimeout)
    If $hLORASR = 0 Then
        ThrowError("Could not finish the calculation.", 3, "RunCalculation", @error)
        KillLORASR()
        SetError(7)
        Return 0
    EndIf

    ; Report sucess
    LogMessage("Calculation complete.", 4, "RunCalculation")
    Return $hLORASR

EndFunc

; Function to cancel failed attempt to load input file
Func CancelLoadInputFile()
    LogMessage("Called `CancelLoadInputFile()`", 5)

    ; Error 1: invalid file name - options "OK" - send "Escape"
    If WinExists($g_sLoadInputFileWindowTitle, "OK") Then
        LogMessage("Invalid filename error. Attempting to cancel.", 3, "CancelLoadInputFile")
        WinActivate($g_sLoadInputFileWindowTitle, "OK")
        Sleep(50)
        Send("{ESCAPE}")
        Sleep(50)
    EndIf

    ; Error 2: file doesn't exist - options "Yes" or "No" to creating new file - send "N" for "No"
    If WinExists($g_sLoadInputFileWindowTitle, "&No") Then
        LogMessage("File doesn't exist error. Attempting to cancel.", 3, "CancelLoadInputFile")
        WinActivate($g_sLoadInputFileWindowTitle, "&No")
        Sleep(50)
        Send("n")
        Sleep(50)
    EndIf

    ; Close the Load Input File dialog box
    WinClose($g_sLoadInputFileWindowTitle)
    WinClose($g_sLoadInputFileWindowTitle)
    WinClose($g_sLoadInputFileWindowTitle)

    ; Kill the process if not already closed
    If WinExists($g_sLoadInputFileWindowTitle) Then
        WinKill($g_sLoadInputFileWindowTitle)
        WinKill($g_sLoadInputFileWindowTitle)
        WinKill($g_sLoadInputFileWindowTitle)
    EndIf

    ; If the window is still there, we haven't succeeded in cancelling
    If WinExists($g_sLoadInputFileWindowTitle) Then
        ThrowError("Cannot close load input file dialog.", 2, "CancelLoadInputFile")
        SetError(2)
        Return 0
    Else
        LogMessage("Successfully closed load input file dialog.", 5, "CancelLoadInputFile")
    EndIf

    ; Exit
    Return (Not @error)

EndFunc

; Function to cancel a crashed program
Func FindError()
    LogMessage("Called `FindError()`", 5)

    ; Declarations
    Local $hError = 0
    Local $bFoundError = False

    ; Check if error window exists
    LogMessage("Checking for error", 5, "FindError")
    $hError = WinActivate("LORASR.exe", "stopped working")

    ; If found, try to exit the error message
    If $hError > 0 Then
        $bFoundError = True
        LogMessage("Error message encountered. Trying to close error message box", 3, "FindError")
        ; Send keystrokes to close window
        Sleep(500)
        Send("{ENTER}")
        Sleep(500)
        Send("{ESCAPE}")
        Sleep(500)
        Send("!c") ; Alt-C
        Sleep(500)
    EndIf

    ; Exit
    Return $bFoundError

EndFunc

; Function to activate and wait for a window, while catching an error state
Func SafeActivate($sWindowTitle, $sWindowText = "", $iWaitTimeout = 10)
    LogMessage("Called `SafeActivate($sWindowTitle = " & $sWindowTitle & ", $sWindowText = " & $sWindowText & ", $iWaitTimeout = " & $iWaitTimeout & ")`", 5)

    ; Declarations
    Local $hWindow = 0
    Local $iLoopCount = 0
    Local $tLoopTimer = TimerInit()

    ; Keep trying until success or failure
    LogMessage("Activating window `" & $sWindowTitle & "`", 5, "SafeActivate")
    While 1
        $iLoopCount += 1
        LogMessage("Loop " & $iLoopCount, 5, "SafeActivate")
        LogMessage("Calling `WinActivate(" & $sWindowTitle & ", " & $sWindowText & ")`", 5, "SafeActivate")
        ; Try to activate requested window
        $hWindow = WinActivate($sWindowTitle, $sWindowText)
        LogMessage("Window handle: " & $hWindow, 5, "SafeActivate")
        If $hWindow > 0 Then
            ; Successfully activated
            LogMessage("Done.", 5, "SafeActivate")
            Return $hWindow
        Else
            LogMessage("Not done...", 5, "SafeActivate")
            ; Check for error
            If FindError() Then
                ThrowError("LORASR gave an error.", 3, "SafeActivate", @error)
                SetError(1)
                Return 0
            EndIf
        EndIf
        ; If no success or failure, wait a short while and then try again
        LogMessage("Sleeping", 5, "SafeActivate")
        Sleep(10)
        LogMessage("Waking", 5, "SafeActivate")
        ; Check for timeout
        If TimerDiff($tLoopTimer) > $iWaitTimeout * 1000 Then
            ThrowError("Timeout in `SafeActivate(" & $sWindowTitle & ", " & $sWindowText & ")`", 3, "SafeActivate", @error)
            SetError(2)
            Return 0
        EndIf
    WEnd

    ; Shouldn't reach this point
    Return 0

EndFunc   ;==>SafeActivate

; Function to force LORASR.exe to quit
Func KillLORASR()
    LogMessage("Called `KillLORASR()`", 5)

    ; Declarations
    Local $iTries = 0, $iMaxTries = 10

    ; Try to close LORASR
    LogMessage("Exiting LORASR", 3, "KillLORASR")
    WinClose($g_sMainWindowTitle)

    ; Try to force LORASR to close
    While WinExists($g_sGenericWindowTitle) And $iTries < $iMaxTries
        LogMessage("Standard exit failed, trying to force quit", 3, "KillLORASR")
        ; May be stuck on loading the input file
        CancelLoadInputFile()
        ; May be stuck with a crash warning
        FindError()
        ; Close all windows
        WinClose($g_sLoadInputFileWindowTitle)
        WinClose($g_sLoadInputFileWindowTitle)
        WinClose($g_sConsoleWindowTitle)
        WinClose($g_sConsoleWindowTitle)
        WinClose($g_sMainWindowTitle)
        WinClose($g_sMainWindowTitle)
        WinClose($g_sGenericWindowTitle)
        WinClose($g_sGenericWindowTitle)
        ; Kill any processes that are not responding
        WinKill($g_sLoadInputFileWindowTitle)
        WinKill($g_sLoadInputFileWindowTitle)
        WinKill($g_sConsoleWindowTitle)
        WinKill($g_sConsoleWindowTitle)
        WinKill($g_sMainWindowTitle)
        WinKill($g_sMainWindowTitle)
        WinKill($g_sGenericWindowTitle)
        WinKill($g_sGenericWindowTitle)
        ; Wait a bit
        Sleep(100)
        $iTries += 1
    WEnd

    ; Report results
    Switch $iTries
        Case 0
            LogMessage("LORASR process closed successfully.", 4, "KillLORASR")
        Case 1 To ($iMaxTries - 1)
            LogMessage("LORASR process forced to close.", 3, "KillLORASR")
        Case Else
            ThrowError("Could not kill the LORASR process.", 2, "KillLORASR", @error)
            SetError(2)
            Return 0
    EndSwitch

    ; Exit
    Return (Not @error)

EndFunc
