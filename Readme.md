# runLORASR

This code comprises six different programs for working with _LORASR_:

* `runLORASR:`
    Run _LORASR_ for a given input file.
* `plotLORASR:`
    Copy data from _LORASR_ output files to _Excel_ plotting spreadsheet.
* `collateLORASR:`
    Collect the transmission data from _LORASR_ output files to a results file.
* `tidyLORASR:`
    Tidy up files from a batch run of _LORASR._
* `sweepLORASR:`
    Sweep through sets of parameters and create individual input files
    for each set of values.
* `batchLORASR:`
    Work through a batch of input files and Run _LORASR_ for each one.

The programs can be run separately, or, for a batch run,
the program `batchLORASR` will work through all the required
tasks for the simulation for all input files in the batch,
including the parameter sweep, running the simulation, plotting
and collating the results, and tidying up the generated files.


# Quick start

### Single run

1. Copy all the following files to a single folder:
    1. `runLORASR.exe`
    2. `plotLORASR.exe`
    3. `tidyLORASR.exe`
    4. The master plot spreadsheet `Plots.xlsx`
    5. The original _LORASR_ program `LORASR.exe`
    6. Your input file(s)
2. Double-click `runLORASR` to start the process
3. Enter the name of the input file
4. _LORASR_ will run automatically
5. Double-click `plotLORASR` to generate the plot spreadsheet
6. Repeat steps 2-5 for each input file
7. Double-click `tidyLORASR` to clean up the output files

### Batch run

1. Copy all the following files to a single folder:
    1. `batchLORASR.exe`
    4. The master plot spreadsheet `Plots.xlsx`
    5. The original _LORASR_ program `LORASR.exe`
    6. Your input files, named `*.in`
2. Double-click `batchLORASR` to automatically run the whole process
    for all input files

### Parametric sweep

1. Copy all the following files to a single folder:
    1. `batchLORASR.exe`
    2. The master plot spreadsheet `Plots.xlsx`
    3. The parameter sweep definition spreadsheet `Sweep.xls`
    4. The original _LORASR_ program `LORASR.exe`
    5. Your input file
2. Modify `Sweep.xls` to contain the parameters you want to sweep and the
    values to sweep through
3. Rename your input file as `Template.txt`
4. Modify the template file so that each parameter that should be swept
    is listed as a variable instead of a value,
    e.g. `POL.=1, DRIFTL.= 3.0, POLEL.= PARAM1, FIELDST.= PARAM2`
5. Double-click `batchLORASR` to automatically run the whole process
    for all swept values of all parameters


# Contents

The executable programs:

* `runLORASR.exe`
* `plotLORASR.exe`
* `collateLORASR.exe`
* `tidyLORASR.exe`
* `sweepLORASR.exe`
* `batchLORASR.exe`

The settings file:

* `runLORASR.ini`

Template files:

* `Plots.xlxs`
* `Sweep.xlsx`
* `Template.txt`

Documentation:

* This Readme file
* Code documentation


# Using the code

Each of the programs can be run by double-clicking in Windows or from the
command line by typing the program name. The programs can either be installed
in the same folder as the _LORASR_ program itself, or they can be run in any
folder without requiring installation.

The settings used by the code are stored in the `runLORASR.ini` settings file.
Most of these settings can be left as the default, and in fact, if the
settings file is not present, the default settings are loaded automatically.

The programs must be able to find the _LORASR_ program (i.e. `LORASR.exe`)
itself. There are two ways to do this: either copy `LORASR.exe` to the same
folder that the `runLORASR` programs are in, or modify `runLORASR.ini` to
tell the code where the _LORASR_ executable is installed.

To use the plotting function, you must have _Microsoft Excel_ installed, and
the master plot file `Plots.xlsx` should either be in the working folder for
the simulation, or in the same folder where the _LORASR_ program is installed.


## batchLORASR

Work through a batch of input files and call `RunLORASR` for each one.

It can use given input files or automatically build input files based on
a parameter sweep definition file and a template input file.
It produces a set of plots for each run, and saves batch results to a
combined file for comparison.

This is the main program that users will interact with. The individual steps
can be run manually by running the other programs, but `batchLORASR` runs all
of these steps automatically.

### Usage

Double-clicking the program or running `batchLORASR` from the command line will
run through the whole batch without requiring further input from the user.

All the program requires are the input file(s) and the location of the _LORASR_
executable. Either copy `LORASR.exe` to the working folder, or modify the
`runLORASR.ini` settings file with the correct installation location.

To set up a parametric sweep, see the usage notes for `sweepLORASR` below.

Note that if the sweep files are present as well as input files, the program
will assume that you want to do a parametric sweep, and will ignore the other
input files. If you just want to use input files, move the `Sweep.xlsx` file
out of the working folder.

If you want to generate output plot spreadsheets or each run, you must have
_Microsoft Excel_ installed, and the master plot file `Plots.xlsx` should
either be in the working folder for the simulation, or in the same folder
where the _LORASR_ program is installed.

### Program method

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


## runLORASR

Run _LORASR_ for a given filename.

It handles the running of the _LORASR_ file, using keyboard shortcuts and window
controls to tell the program what to do. Intended to be used as part of a batch,
but can run on its own, either from the command line with the input file as a
parameter, or interactively where the user enters the input filename.

### Usage

Double-clicking the program or running `runLORASR` from the command line will
prompt the user to enter the input file name. The input file name can also be
given to the program as a command line parameter, e.g.

    runLORASR dtl.in

All the program requires is the input file and the location of the _LORASR_
executable. Either copy `LORASR.exe` to the working folder, or modify the
`runLORASR.ini` settings file with the correct installation location.

### Program method

1. Get input file name
2. Check input and program files exist
3. Start the _LORASR_ software
4. Open the input file
5. Run the calculation
6. Wait for the calculation window to be complete
7. Quit the program


## plotLORASR

Reads in the output data produced by _LORASR_ in various text files and creates
a results spreadsheet, based on a template with plots already defined.
Produces all the same standard plots produced by _LORASR._

This program can be run by itself in a folder where _LORASR_ has been run
manually, or can be called for each run of a batch process.

Note that running _LORASR_ overwrites the output files, so in a batch run
_PlotLORASR_ must be called before starting the next run in the batch.

### Usage

After running _LORASR_ manually, or using the `runLORASR.exe` program to run
_LORASR_ automatically, there will be a number of output files in the working
directory, such as `bucent` and `transenv`.

Double-clicking the program or running `plotLORASR` from the command line will
load all the data from these output files into a spreadsheet file.

You must have _Microsoft Excel_ installed, and the master plot file `Plots.xlsx`
should either be in the working folder for the simulation, or in the same folder
where the _LORASR_ program is installed.

### Program method

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

Data type for step 5 are:

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


## sweepLORASR

Sweep through sets of parameters and create individual input files
for each set of values.

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
using `batchLORASR.exe`

### Usage

Double-clicking the program or running `sweepLORASR` from the command line will
load all the parametric sweep definitions and create all the individual input
files for each combination of parameter values.

Normally `sweepLORASR` wouldn't be run by itself, but rather run as part of a
batch. To process the batch, set up the template and parameter sweep spreadsheet
as below, then double-click `batchLORASR.exe` or type `batchLORASR` at the
command prompt.

The simulation must be specified in a template file. This is an input file
where the parameters you want to change have been replaced by variable names
such as `PARAM1`, `PARAM2` etc.

For example, to sweep through a number of possible values for the field length
and field gradient of a quadrupole, you might modify one line in the input file
from

    POL.=1, DRIFTL.= 3.0, POLEL.= 7.3, FIELDST.= 5040

to

    POL.=1, DRIFTL.= 3.0, POLEL.= PARAM1, FIELDST.= PARAM2

The default name for the template file is `Template.txt`. This can be changed
in the `runLORASR.ini` settings file.

Once the template file is ready, the parameters to sweep and the values to use
must be specified in a parameter sweep definition spreadsheet. The default name
for the definition file is `Sweep.xlsx`, which can also be changed in the
`runLORASR.ini` settings file.

In the first sheet in the spreadsheet, put in the name and units for each
parameter that you want to sweep through. The number of parameters and number
of values for each parameter will be calculated automatically, so these don't
need to be manually entered.

In the second sheet, there will be a column for each parameter that you
entered on the first sheet. Fill in the values to be simulated for each
parameter.

### Program method

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

## collateLORASR

Collate transmission and beam core results from a batch run of _LORASR_ into
a batch results file.

This is normally handled automatically by `batchLORASR`, but there may be cases
where a long batch was interrupted and the batch results are missing.
`collateLORASR` works through a set of input and output files, and collates
all the results together.

### Usage

The working folder should contain the various input files `*.in` and
corresponding run files `*.out`, either all together in one folder or with
separate input and run subfolders. The default names for the these subfolders
are `Input\` and `Runs\`, although these can be changed in the `runLORASR.ini`
settings file.

Double-clicking the program or running `collateLORASR` from the command line
will sort through these files and produce a batch results output file. The
default name for the output file is `Batch results.csv`, which can also be
changed in the `runLORASR.ini` settings file.

### Program method

1. Get a list of all output files in the current folder or the Runs subfolder
2. Create a blank output file
3. For each unique run:
    1. Find the input file `*.in`
    2. Find the _latest_ run file `*.out`
    3. Load the starting number of macroparticles from the input file
    4. Load the final number of macroparticles from the output file
    5. Calculate the transmission
    6. Load the final proportion of particles in core from the output file
    7. Save these results to the output file


## tidyLORASR

Tidy up files from a batch run of _LORASR._

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
`runLORASR.ini` settings file.

### Usage

Double-clicking the program or running `tidyLORASR` from the command line will
move files into their correct subfolders and delete extraneous temporary files.

The subfolder names can be specified in the `runLORASR.ini` settings file.

### Program method

1. For each run
    (either called by the batch program or loop through all output files):
    1. Determine whether the run was successful or not
    2. Move the files for that run to the correct place:
        1. All files for unsuccessful runs go to the subfolder
            defined by the `IncompleteFolder` setting
        2. Input files for successful runs go to the subfolder
            defined by the `InputFolder` setting
        3. Run files for successful runs go to the subfolder
            defined by the `RunFolder` setting
        4. Output files for successful runs go to the subfolder
            defined by the `OutputFolder` setting
2. Also tidy up batch-related files:
    1. Any files missed by the run-by-run tidying are tidied as above
    2. Master plot file goes to the subfolder defined by `$sInputFolder`
    3. Data output files produced directly by _LORASR_ are deleted,
        as they are no longer required
    4. _LORASR_ executable file is deleted from the working directory
        (this shouldn't be the master version)
    5. Any files marked as "old" are also deleted


# Example settings file

    [Files and folders]
    ProgramPath="C:\Program Files (x86)\LORASR"
    SimulationProgram="LORASR.exe"
    SweepFile="Sweep.xlsx"
    TemplateFile="Template.txt"
    ResultsFile="Batch results.csv"
    PlotFile="Plots.xlsx"
    LogFile="runLORASR.log"
    InputFolder="Input"
    OutputFolder="Output"
    RunFolder="Runs"
    IncompleteFolder="Incomplete"

    [Options]
    Cleanup=True
    ConsoleVerbosity=5
    LogFileVerbosity=3
    MessageVerbosity=1


# Readme details

* Author: Matt Easton matt.easton@pku.edu.cn
* Last modified: 2017-09-15
* Relates to _runLORASR_ code version: v0.4
