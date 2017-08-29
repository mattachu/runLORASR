# runLORASR code documentation
* Author: Matt Easton
* Created: 2017-08-29
* Code version: v0.4.0

-------------------------------------------------------------------------------

# runLORASR

## Function
Run _LORASR_ for a given filename.

## Description
A script to run the actual beam dynamics simulation for a given input file.
It handles the running of the _LORASR_ file, using keyboard shortcuts and window
controls to tell the program what to do. Intended to be used as part of a batch,
but can run on its own, either from the command line with the input file as a
parameter, or interactively where the user enters the input filename.

## Method
1. Get input file name
2. Check input and program files exist
3. Start the _LORASR_ software
4. Open the input file
5. Run the calculation
6. Wait for the calculation window to be complete
7. Quit the program

## Versions
* 0.1: First version
* 0.2: Adjusted wait handling to monitor calculation progress and only
    continue when complete; better error handling
* 0.4: Major rewrite to use modules and functions. Split into two parts:
    * _runLORASR.Run_ is a reusable library that holds all the
        actual logical code
    * _runLORASR_ is a wrapper to run as a stand-alone executable.

## Functions
Note that global variables at the top of _runLORASR.Run_ are used to specify
the names of the program windows in _LORASR._ These are used by the script in
interacting with the simulation program.

### SetupLORASR
Preparation of the simulation environment (copies files to the correct place).

Function call:

    SetupLORASR($sWorkingDirectory, $sProgramPath, $sSimulationProgram)

Parameters:
* `$sWorkingDirectory:` Working directory in which the simulation will be run
* `$sProgramPath`: Path to find the program files and executables
* `$sSimulationProgram`: Filename of the simulation program executable

Returns: Full path and file name of the simulation program executable

### RunLORASR
Main function that runs through the method above.

Function call:

    RunLORASR($sRun, $sWorkingDirectory, $sSimulationProgramPath, $sInputFolder)

Parameters:
* `$sRun`: Name for this particular run (used in batch operation)
* `$sWorkingDirectory`: Working directory in which the simulation will be run
* `$sSimulationProgramPath`: Full path and file name of the simulation program
    executable
* `$sInputFolder`: Folder in which to check for input files.

Returns: success or failure

### LoadInputFile
Handles the loading of the input file.

Function call:

    LoadInputFile($sInputFile)

Parameters:
* `$sInputFile`: Full path and file name of input file to load.

Returns: handle of the Load Input File dialog window

### RunCalculation
Tells _LORASR_ to run the calculation and waits for the result.

Function call:

    RunCalculation()

Parameters: none

Returns: handle of the completed calculation window

### CancelLoadInputFile
Handles any errors that arise during the loading of the input file.

Function call:

    CancelLoadInputFile()

Parameters: none

Returns: success or failure

### FindError
Checks whether _LORASR_ is stuck with an error message, and tries to clear
the error message if it finds one.

Function call:

    FindError()

Parameters: none

Returns: whether not an error was found

### SafeActivate
Wrapper around WinActivate that tries a number of times to activate the given
window. It checks for any errors during the process and tries to clear up any
problems found and try again.

Function call:

    SafeActivate($sWindowTitle, $sWindowText, $iWaitTimeout)

Parameters:
* `$sWindowTitle`: Matches the title of the window to activate
* `$sWindowText`: If given, only matches windows with this text as well as the
    matching title
* `$iWaitTimeout`: How long to keep trying for (in seconds) - default 10s

Returns: handle of the activate window

### KillLORASR
Attempt to close _LORASR_ gracefully, and finally force it to quit
if unresponsive.

Function call:

    KillLORASR()

Parameters: none

Returns: success or failure

--------------------------------------------------------------------------------

# batchLORASR

## Function
Work through a batch of input files and call RunLORASR for each one.

## Description
The wrapper script that handles batch processing of _LORASR_ simulations.
This program is the one the user interacts with, and it calls the other
programs as it works through the bunch. It can use given input files or
automatically build input files based on a parameter sweep definition file
and a template input file. It produces a set of plots for each run, and saves
batch results to a combined file for comparison.

## Method
1. Get working directory
2. Copy input files to working directory if required
3. Run sweep module to automatically produce input files if parameter
    sweep definition file is present
4. Create a batch results output file to store the summarised results
5. Get list of input files in working directory
6. Run for each input file sequentially:
    1. Run LORASR
    2. Create plots spreadsheet for this run
    3. Write results to combined batch results file
    4. Tidy up input and output files for this run
    5. Next input file
7. Tidy up batch files

## Versions
* 0.1: First version
* 0.2: Add ability to build multiple input files by sweeping parameters
* 0.3: Allow standalone running based on executables in working directory;
    better sweep handling; better error handling
    * 0.3.1: Error handling if no output file exists
* 0.4: Major rewrite to use modules and functions. Split into two parts:
    * _runLORASR.Batch_ is a reusable library that holds all the
        actual logical code
    * batchLORASR is a wrapper to run as a stand-alone executable.

## Functions

### BatchLORASR
Main function that steps through the method above and calls the subfunctions
below and the other modules.

Function call:

    BatchLORASR($sWorkingDirectory, $sProgramPath, $sSimulationProgram,
        $sSweepFile, $sTemplateFile, $sResultsFile, $sPlotFile, $sInputFolder,
        $sOutputFolder, $sRunFolder, $sIncompleteFolder, $bCleanup)

Parameters:
* `$sWorkingDirectory`: Folder in which to run the batch
* `$sProgramPath`: Folder that holds to simulation program files
* `$sSimulationProgram`: Filename of simulation program executable
* `$sSweepFile`: Filename of parameter sweep definition spreadsheet
* `$sTemplateFile`: Filename of template input file from which to build input
    files for a parameter sweep
* `$sResultsFile`: Filename of the batch results file to be created
* `$sPlotFile`: Filename of the master plot file from which to build the plot
    workbooks for each run
* `$sInputFolder`: Subfolder that holds input files
    (if not in the working directory)
* `$sOutputFolder`: Subfolder in which to store output files when tidying up
* `$sRunFolder`: Subfolder in which to store simulation run log files
    when tidying up
* `$sIncompleteFolder`: Subfolder in which to store incomplete runs
    when tidying up
* `$bCleanup`:  Whether or not to tidy up files into subfolders and delete
    leftover files at the end of the batch

Returns: success or failure

### FindInputFiles
Find input files in working directory or copy from input subfolder if required.

Function call:

    FindInputFiles($sWorkingDirectory, $sInputFolder, $sSweepFile,
        $sTemplateFile, $sPlotFile, $sProgramPath)

Parameters:
* `$sWorkingDirectory`: Folder in which batch is running
* `$sInputFolder`: Subfolder in which to look for input files
    (in addition to checking the working directory)
* `$sSweepFile`: Filename of parameter sweep definition spreadsheet to be found
* `$sTemplateFile`: Filename of parameter sweep template input file to be found
* `$sPlotFile`: Filename of the master plot spreadsheet to be found
    (searches working directory first, then input subfolder, then program folder)
* `$sProgramPath`: Path to the _LORASR_ program folder, where the main master
    plot workbook should be found.

Returns: success or failure

### RunSweepLORASR
Searches for parameter sweep definition files, and calls SweepLORASR if found.

Function call:

    RunSweepLORASR($sWorkingDirectory, $sSweepFile, $sTemplateFile,
        $sResultsFile, $sInputFolder)

Parameters:
* `$sWorkingDirectory`: Folder in which the batch is running
* `$sSweepFile`: Filename of parameter sweep definition spreadsheet to be found
* `$sTemplateFile`: Filename of parameter sweep template input file to be found
* `$sResultsFile`: Filename of batch results output file to create
* `$sInputFolder`: Subfolder in which to search for input files,
    in addition to working directory

Returns: whether or not parameter sweep input files were set up

### SaveResults
Loads details of the current run from the simulation run log file, and writes
main results to batch results summary output file.

Function call:

    SaveResults($sRun, $sResultsFile, $sWorkingDirectory)

Parameters:
* `$sRun`: Name of the current run, either a batch identifier or a list of
    parameter values from the parameter sweep
* `$sResultsFile`: Filename of the output file to which to write out
    batch results
* `$sWorkingDirectory`: Folder in which batch is running

Returns: success or failure

-------------------------------------------------------------------------------

# plotLORASR

## Function
Copy data from _LORASR_ output files to _Excel_ plotting spreadsheet.

## Description
Reads in the output data produced by _LORASR_ in various text files and creates
a results spreadsheet, based on a template with plots already defined.
Gives all the standard plots produced by _LORASR._
This can be run by itself in a folder where _LORASR_ has been run manually,
or can be called for each run of a batch process.
Note that running _LORASR_ overwrites the output files, so in a batch run
_PlotLORASR_ must be called before starting the next run in the batch.

## Method
1. Find master spreadsheet
2. Create plot spreadsheet from master
3. Open Excel
4. Open the spreadsheet
5. For each type of data:
    1. Load data from file to array
    2. Write data to correct place in spreadsheet
    3. Next data type
6. Save the spreadsheet
7. Close Excel

Data type are:
* Transverse envelope data
* Longitudinal envelope data
* Input x-emittance data
* Output x-emittance data
* Input y-emittance data
* Output y-emittance data
* Input z-emittance data
* Output z-emittance data
* Bunch centre data
* Emittance growth data
* Phase advance data
* Emittance values

## Versions
* 0.1: First version
* 0.2: Allow standalone running using master plot file in working directory;
    better error handling
* 0.4: Major rewrite to use modules and functions. Split into two parts:
    * _runLORASR.Plots_ is a reusable library that holds all the
        actual logical code
    * _plotLORASR_ is a wrapper to run as a stand-alone executable.


## Functions

### GetFileSettings
Defines input and output settings for each data type;
changes to the _LORASR_ code may require changes here.

Function call:

    GetFileSettings($sDataType, ...)

Parameters:
* `$sDataType`: which type of data is currently being proceesed
* All other parameters are  (using `ByRef`) as a method of passing results
    back to the calling function.

Returns: success or failure is the main result, and the file settings
    are returned by the referenced parameters

### PlotLORASR
Main function that works through the method above and calls the other functions.

Function call:

    PlotLORASR($sWorkingDirectory, $sPlotFile, $sMasterPlotFile, $sMasterPath)

Parameters:
* `$sWorkingDirectory`: folder that holds the output files from _LORASR_ and
    in which the plot file will be saved
* `$sPlotFile`: name of the plot file to be created
* `$sMasterPlotFile`: name of the master plot file from which the plots for
    the current run will be generated
* `$sMasterPath`: file location where the master plot file is stored

Returns: success or failure

### PlotAllData
Wrapper function to work through each data type and plot data into the workbook
for the current run.

Function call:

    PlotAllData(ByRef $oWorkbook, $sWorkingDirectory)

Parameters:
* `$oWorkbook`: handle of the workbook currently being processed
* `$sWorkingDirectory`: location of the folder that contains the output files
    from _LORASR_

Returns: success or failure

### PlotData
Function to plot a given set of data (from one or sometimes two _LORASR_
output files) into the given Excel workbook (usually one sheet for the given
type of data).

Function call:

    PlotData($sDataType, ByRef $oWorkbook, $sWorkingDirectory)

Parameters:
* `$sDataType`: the particular type of data to be plotted at this stage
    (uses GetFileSettings to find the various required settings for processing
     this particular data type)
* `$oWorkbook`: handle of the workbook currently being processed
* `$sWorkingDirectory`: location of the folder that contains the output files
    from LORASR

Returns: success or failure

### CreatePlotSpreadsheet
Function to create a new spreadsheet from the master.

Function call:

    CreatePlotSpreadsheet($sWorkingDirectory, $sPlotFile, $sMasterPlotFile,
        $sMasterPath)

Parameters:
* `$sWorkingDirectory`: folder in which the plot file will be saved
* `$sPlotFile`: name of the plot file to be created
* `$sMasterPlotFile`: name of the master plot file from which the plots for
    the current run will be generated
* `$sMasterPath`: file location where the master plot file is stored

Returns: full path and filename of the newly created file (as a string)

### SaveAndClosePlotSpreadsheet
Does what the name suggests: saves the workbook and quits Excel (based on
object handles of the application and workbook).

Function call:

    SaveAndClosePlotSpreadsheet(ByRef $oExcel, ByRef $oWorkbook)

Parameters:
* `$oExcel`: handle for the Excel application object,
    as created by the code when Excel was started
* `$oWorkbook`: handle for the Excel workbook object,
    as created by the code when the workbook was opened

Returns: success or failure

### LoadData
Function to load specific data from a file into an array.

Function call:

    LoadData($sDataFile, $sWorkingDirectory, $iDataStart, $iDataEnd)

Parameters:
* `$sDataFile`: filename of file containing the data to be loaded
* `$sWorkingDirectory`: file location of file containing the data to be loaded
* `$iDataStart`: starting row within the file of the particular data to load
* `$iDataEnd`: starting row within the file of the particular data to load

Returns: an array of strings containing the requested data

### WriteData
Function to write the given data to a specific location in the spreadsheet.

Function call:

    WriteData(ByRef $oWorkbook, $iWorksheet, $sDataLocation, $asData)

Parameters:
* `$oWorkbook`: handle of the workbook currently being processed
* `$iWorksheet`: the sheet within the workbook
    into which the current data should be written
* `$sDataLocation`: cell reference within the given worksheet
    into which the current data should be written
* `$asData`: the array of data that should be written

Returns: success or failure

--------------------------------------------------------------------------------

# sweepLORASR

## Function
Sweep through sets of parameters and create individual input files
for each set of values.

## Description
By defining which parameters should be swept, what values should be used
for each sweep, and where these parameters fit into the overall simulation,
this code allows an arbitrary number of simulations to be run with various
different values of various parameters.
The parameters and their values are specified in a spreadsheet, which also
estimates the batch run time for the sweep.
The simulation is specified in a template file, where placeholders for the
parameters to be varied are automatically replaced by the code to create
a set of multiple input files from this template, one input file for each
combination of parameter values.
The resulting input files can be simulated manually, or automatically
using _batchLORASR._

## Method
1. Get working directory
2. Clear up existing input files in working directory
3. Copy parameter sweep files to working directory if required
4. Check whether parameter sweep files exists
5. Get parameter details from sweep definition file
    1. Open _Excel_
    2. Open the spreadsheet
    3. Read in number of parameters
    4. Read in definition of each parameter from first sheet
    5. Read in parameter values from second sheet
    6. Close _Excel_
6. Copy the template to one input file for each of the values of the
   first parameter:
    1. Build filename
    2. Copy files
    3. Replace the dummy code for this first parameter with the
       value of the parameter
    4. Next value of first parameter
7. For each successive parameter, copy all existing input files once for
   each value of that parameter:
    1. Find existing files
    2. For each value of the current parameter:
        1. For each existing input file:
            1. Build new filename for this input file and this set of
               parameter values
            2. Copy to new input file
            3. Replace the dummy code for this current parameter with the
               value of the parameter
            4. Next file
        2. Next value of current parameter
    3. Delete previous base files (as we are now a level deeper)
    4. Next parameter
8. Prepare batch results file with headings based on parameter sweep definition

## Versions
* 0.1: First version
* 0.2: Build headers for batch results csv file with correct headings from the
    sweep definition file
    * 0.2.1: Fix for sweeps of a single parameter
* 0.4: Major rewrite to use modules and functions. Split into two parts:
    * _runLORASR.Sweep_ is a reusable library that holds all the
        actual logical code
    * _sweepLORASR_ is a wrapper to run as a stand-alone executable.

## Functions

### SweepLORASR
Main function that runs through the method above and calls the utility
functions below.

Function call:

    SweepLORASR($sWorkingDirectory, $sSweepFile, $sTemplateFile, $sResultsFile,
        $sInputFolder)

Parameters:
* `$sWorkingDirectory`: location of the folder that should be used to prepare
    the parameter sweep
* `$sSweepFile`: name of the file that specifies the parameters to be swept
* `$sTemplateFile`: name of the file that specifies the rest of the simulation
    apart from the swept parameters
* `$sResultsFile`: name of the file that should be created to contain the
    results of the parameters sweep
* `$sInputFolder`: name of a subfolder within the working directory where the
    input files may be found

Returns: success or failure

### DeleteInputFiles
Function to clear up existing input files in the working directory,
otherwise these would confuse the batch run. Incomplete batch runs may leave
multiple input files that have not been processed properly.

Function call:

    DeleteInputFiles($sWorkingDirectory , $sTemplateFile)

Parameters:
* `$sWorkingDirectory`: location of the folder that should be used to prepare
    the parameter sweep
* `$sTemplateFile`: name of the file that specifies the current simulation
    (this is the only input file that will not be deleted)

Returns: success or failure

### FindSweepFiles
Function to find parameter sweep definition and template files and copy them
to the working directory.

Function call:

    FindSweepFiles($sWorkingDirectory, $sSweepFile, $sTemplateFile,
        $sInputFolder)

Parameters:
* `$sWorkingDirectory`: location of the folder that should be used to
    prepare the parameter sweep
* `$sSweepFile`: name of the parameter sweep definition file to search for
* `$sTemplateFile`: name of the simulation template input file to search for
* `$sInputFolder`: name of a subfolder within the working directory where
    these files may be found

Returns: success or failure

### LoadSweepParameters
Function to load the parameters and values from the sweep definition
spreadsheet into two arrays.

Function call:

    LoadSweepParameters(ByRef $asParameters, ByRef $asValues,
        $sWorkingDirectory, $sSweepFile)

Parameters:
* `$asParameters`: this referenced array parameter is overwritten with the
    set of sweep parameters
* `$asValues`: this referenced array parameter is overwritten with the
    set of sweep parameter values
* `$sWorkingDirectory`: location of the folder that should be used to
    prepare the parameter sweep
* `$sSweepFile`: name of the file that specifies the parameters to be swept

Returns: success or failure as the main result, and the parameters and their
    values by referenced array parameters

### SweepParameter
Function to sweep through a parameter and create input files for each value
of that parameter.
Called recursively to work through all parameters and all possible parameter
value combinations.
For the first parameter, this copies from the template file, one input file
per parameter value.
For subsequent parameters, this copies from the existing input files that have
just been created, so for each recursion the number of input files is multiplied.

Function call:

    SweepParameter($asParameters, $asValues, $iParameter, $sWorkingDirectory,
        $sTemplateFile)

Parameters:
* `$asParameters`: set of parameters to be swept
* `$asValues`: set of parameter values to sweep over
* `$iParameter`: the index of the parameter currently being processed
    (zero-indexed)
* `$sWorkingDirectory`: location of the folder that should be used to prepare
    the parameter sweep
* `$sTemplateFile`: name of the file that specifies the rest of the simulation
    apart from the swept parameters

Returns: success or failure

### CreateResultsFile
Function to create a results file in which to store the summary results for
each set of values in the parameter sweep. This is included here as at this
point in a batch run, the set of parameters and their values are loaded in
memory, and can easily be written out as headers to the results file.

Function call:

    CreateResultsFile($asParameters, $sWorkingDirectory, $sResultsFile)

Parameters:
* `$asParameters`: set of parameters to be swept
    (used to define headers in the output file)
* `$sWorkingDirectory`: location of the folder that should be used to
    prepare the parameter sweep
* `$sResultsFile`: name of the file that should be created to contain the
    results of the parameters sweep

Returns: success or failure

--------------------------------------------------------------------------------

# tidyLORASR

## Function
Tidy up files from a batch run of _LORASR._

## Description
As batch files can produce a lot of input files, output files, results files,
log files, and temporary data files, without some kind of handling of these
files it is hard to find the information required. Also, leaving behind various
output files can make re-running a batch more complicated.
This program puts input files in an input folder, output files in an output
folder, run data files in a run folder and also deletes old files that are
no longer needed. If a run was unsuccessful, all related files are stored in a
separate folder for incomplete runs, so they can be easily rerun by the user.
Subfolders are created if they don't exist already.

Files that pertain to the whole batch are left in the parent folder:
the parameter sweep definition spreadsheet and template input file
(if using a parameter sweep), the overall log file, and the batch results
summary file.

This part of a batch can be bypassed (i.e. don't tidy any files, keep them
where they are) by setting the option `Cleanup` to `FALSE` in the
_runLORASR.ini_ settings file.

When called as part of a batch, the main routines are called for the current
run at the end of processing, before moving onto the next run in the batch,
and then the final cleanup routine is called at the end of the batch.
When run as a standalone program, it checks the status of each run in the
working directory and tidies up accordingly, before calling the final cleanup
routine itself.

## Method
1. For each run
    (either called by the batch program or loop through all output files):
    1. Determine whether the run was successful or not
    2. Move the files for that run to the correct place:
        1. All files for unsuccessful runs go to the subfolder
            defined by `$sIncompleteFolder`
        2. Input files for successful runs go to the subfolder
            defined by `$sInputFolder`
        3. Run files for successful runs go to the subfolder
            defined by `$sRunFolder`
        4. Output files for successful runs go to the subfolder
            defined by `$sOutputFolder`
2. Also tidy up batch-related files:
    1. Any files missed by the run-by-run tidying are tidied as above
    2. Master plot file goes to the subfolder defined by `$sInputFolder`
    3. Data output files produced directly by _LORASR_ are deleted,
        as they are no longer required
    4. _LORASR_ executable file is deleted from the working directory
        (this shouldn't be the master version)
    5. Any files marked as "old" are also deleted

## Versions
* 0.0.1: First version didn't get to version 0.1 as the rewrite to
    version 0.4 started at this point
* 0.4: Major rewrite to use modules and functions. Split into two parts:
    * _runLORASR.Tidy_ is a reusable library that holds all the
        actual logical code
    * _tidyLORASR_ is a wrapper to run as a stand-alone executable.

## Functions

### TidyIncompleteRun
Tidy up files from an incomplete run: all files go into the
specified "incomplete" folder.

Function call:

    TidyIncompleteRun($sRun, $sWorkingDirectory, $sIncompleteFolder)

Parameters:
* `$sRun`: identifier (base filename) for the current run
* `$sWorkingDirectory`: folder to be tidied up
* `$sIncompleteFolder`: name of subfolder for incomplete runs

Returns: success or failure

### TidyCompletedRun
Tidy up files from a completed run: files go into the specified folders.

Function call:

    TidyCompletedRun($sRun, $sWorkingDirectory, $sInputFolder, $sOutputFolder,
        $sRunFolder)

Parameters:
* `$sRun`: identifier (base filename) for the current run
* `$sWorkingDirectory`: folder to be tidied up
* `$sInputFolder`: name of subfolder for input files for each run
* `$sOutputFolder`: name of subfolder for output files for each run
* `$sRunFolder`: name of subfolder for _LORASR_ run data output files (.out) for each run

Returns: success or failure

### TidyRunFiles
Test whether a particular run was completed or not, and tidy up accordingly.

Function call:

    TidyRunFiles($sRun, $sWorkingDirectory, $sInputFolder, $sIncompleteFolder,
        $sOutputFolder, $sRunFolder)

Parameters:
* `$sRun`: identifier (base filename) for the current run
* `$sWorkingDirectory`: folder to be tidied up
* `$sInputFolder`: name of subfolder for input files for each run
* `$sIncompleteFolder`: name of subfolder for incomplete runs
* `$sOutputFolder`: name of subfolder for output files for each run
* `$sRunFolder`: name of subfolder for _LORASR_ run data output files (`.out`)
    for each run

Returns: success or failure

### TidyAllRunFiles
Work through all input files, test whether completed or not, and tidy up
accordingly. This is the wrapper called by the executable program _tidyLORASR._
It is not used in a batch process.

Function call:

    TidyAllRunFiles($sWorkingDirectory, $sInputFolder, $sIncompleteFolder,
        $sOutputFolder, $sRunFolder)

Parameters:
* `$sWorkingDirectory`: folder to be tidied up
* `$sInputFolder`: name of subfolder for input files for each run
* `$sIncompleteFolder`: name of subfolder for incomplete runs
* `$sOutputFolder`: name of subfolder for output files for each run
* `$sRunFolder`: name of subfolder for _LORASR_ run data output files (`.out`)
    for each run

Returns: success or failure

### TidyBatchFiles
Tidy up all files used or created by the batch process.
Run after the above functions.

Function call:

    TidyBatchFiles($sWorkingDirectory, $sSimulationProgram, $sSweepFile,
        $sTemplateFile, $sPlotFile, $sInputFolder, $sOutputFolder, $sRunFolder)

Parameters:
* `$sWorkingDirectory`: folder to be tidied up
* `$sSimulationProgram`: filename of _LORASR_ executable
    (gets deleted)
* `$sSweepFile`: filename of parameter sweep definition spreadsheet
    (gets left in main folder)
* `$sTemplateFile`: filename of template input file for parameter sweep
    (gets left in main folder)
* `$sPlotFile`: filename of master spreadsheet file for plotting data
    (gets moved to input subfolder)
* `$sInputFolder`: name of subfolder for input files for each run
* `$sOutputFolder`: name of subfolder for output files for each run
* `$sRunFolder`: name of subfolder for _LORASR_ run data output files (`.out`)
    for each run

Returns: success or failure

### TidySimulationFiles
Tidy up the files produced directly by the _LORASR_ simulation code.
Called by _TidyBatchFiles_ above. If the names of data files created
by _LORASR_ were to change, this is where the references need to be edited.

Function call:

    TidySimulationFiles($sWorkingDirectory)

Parameters:
* `$sWorkingDirectory`: folder to be tidied up

Returns: success or failure

--------------------------------------------------------------------------------

# runLORASR.Functions

## Function
Functions used by _runLORASR._

## Description
Any functions that are common to all modules of the _runLORASR_ code are now
kept in a separate function library, which is included into the other modules.

## Method
There is no logical method followed overall, these are just utility functions.
See below for details.

## Versions
* 0.4: Created as part of a major rewrite of all code to use modules and
    functions.

## Functions

### GetSettings
Function to read settings from _runLORASR.ini_ file.
This allows all modules to access the same settings without requiring
global scope variables.
Once the settings are read from the file, they are passed to other functions
as required.

Function call:

    GetSettings($sWorkingDirectory, ...)

Parameters:
* `$sWorkingDirectory`: the folder in which _runLORASR_ is currently working
* All other parameters are referenced, and will be overwritten with the
    settings loaded from the ini file.

Returns: success or failure as main result, with settings returned via
    referenced parameters

### FindFile
Function to find a specified file or set of files in the working directory
or in a master directory, and also optionally copy the master file into the
working directory.

Function call:

    FindFile($sFindFileName, $sWorkingDir, $sMasterDir, $bCopy)

Parameters:
* `$sFindFileName`: filename of the file(s) to look for - accepts wildcards
* `$sWorkingDir`: the folder in which _runLORASR_ is currently working
* `$sMasterDir`: the folder in which a master version of the file may be found
* `$bCopy`: whether or not to copy the master file to the working folder

Returns: full path to found file

### CopyFiles
Function to copy a file or set of files from one folder to another,
with logging of results.

Function call:

    CopyFiles($sCopyFileName, $sCopySourceFolder, $sCopyDestinationFolder,
        $bOverwrite)

Parameters:
* `$sCopyFileName`: filename of the file(s) to copy - accepts wildcards
* `$sCopySourceFolder`: the folder from which to copy
* `$sCopyDestinationFolder`: the folder into which to copy
* `$bOverwrite`: whether or not to overwrite existing files in the
    destination folder

Returns: success or failure

### MoveFiles
Function to move a file or set of files from one folder to another,
with logging of results.

Function call:

    MoveFiles($sMoveFileName, $sMoveSourceFolder, $sMoveDestinationFolder,
        $bOverwrite)

Parameters:
* `$sMoveFileName`: filename of the file(s) to move - accepts wildcards
* `$sMoveSourceFolder`: the folder from which to move
* `$sMoveDestinationFolder`: the folder into which to move
* `$bOverwrite`: whether or not to overwrite existing files in the
    destination folder

Returns: success or failure

### DeleteFiles
Function to delete a file or set of files from a given folder,
with logging of results.

Function call:

    DeleteFiles($sDeleteFileName, $sSearchFolder)

Parameters:
* `$sDeleteFileName`: filename of the file(s) to delete - accepts wildcards
* `$sSearchFolder`: the folder from which to delete

Returns: success or failure

### LogMessage
Function to log a message to the console (if running interactively), to the log
file, and to show message boxes.
The global settings define what importance level of messages get sent to
what log. For example, with the settings for console set to 5, for log file
set to 3 and for message boxes set to 2, then a level 3 importance message
will be written to the console and the log file, but will not display a
message box. A level 1 importance message will show all three.

Function call:

    LogMessage($sMessageText, $iImportance, $sFunctionName, $sLogFile, $sWorkingDirectory)

Parameters:
* `$sMessageText`: the message to be logged
* `$iImportance`: the importance of the message
    (determines to where the message is output)
* `$sFunctionName`: the function sending the message
* `$sLogFile`: the name of the log file (normally _runLORASR.log_)
* `$sWorkingDirectory`: the folder in which _runLORASR_ is currently working
    (to locate the log file)

Returns: success or failure

### WriteToLogFile
Function to write to the log file, called by _LogMessage_ if required.

Function call:

    WriteToLogFile($sMessageText, $sLogFile, $sWorkingDirectory)

Parameters:
* `$sMessageText`: the message to be logged
* `$sLogFile`: the name of the log file (normally _runLORASR.log_)
* `$sWorkingDirectory`: the folder in which _runLORASR_ is currently working
    (to locate the log file)

Returns: success or failure

### CreateLogFile
Function to create a new log file, called at the start of each _runLORASR_
program.

Function call:

    CreateLogFile($sLogFile, $sWorkingDirectory)

Parameters:
* `$sLogFile`: the name of the log file to be created (normally _runLORASR.log_)
* `$sWorkingDirectory`: the folder in which _runLORASR_ is currently working
    (in which to create the log file)

Returns: success or failure

### ThrowError
Function to handle errors, using the same "importance" scheme as _LogMessage._

Function call:

    ThrowError($sErrorText, $iImportance, $sFunctionName, $iErrorCode)

Parameters:
* `$sErrorText`: the error message to be logged
* `$iImportance`: the importance of the message
    (determines to where the message is output)
* `$sFunctionName`: the function where the error occurred
* `$iErrorCode`: the code given by the program's error handler
    (for troubleshooting)

Returns: success
    (don't want to return an error from an error handler,
     otherwise we get stuck in a loop)
