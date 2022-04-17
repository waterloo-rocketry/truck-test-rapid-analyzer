# truck-test-rapid-analyzer
A quick script for processing DAQ .csv data to excel formats for quick processing of recovery truck test data. That is, it will perform a full parachute aerodynamic analysis and export the results to excel when provided with just timestamps, wind, and force data for multiple files. The original rationale is to allow the quick analysis of multiple parachute configurations, as was necessary for the development of the reefed parachute system. 

## How to use

In order to use it, the following things need to be set up.

 1. The basic configuration variables must be set. These are currently located in the header of the main python file. By reading the comments, the exact functionality can be undertsood and tuned. For the best option, ensure that the FILE_INPUT_HITLIST, CONDENSED_EXPORT_VERSION, and USE_CUSTOM_SHEETNAMES are set to 'True'. This should already be done by default, but looking over the variables is a good idea nontheless.
 2. The location of the csv data must be specified in the FILE_SUBDIRECTORY. This is also the location where the results of the analysis will be stored. It is important to ensure that the csv data is of the right format. An example csv file can be found in tests/example_data_directory
 3. The target_list and target_list_sheetnames must be configured. Every file name (without the file extension, this is set in the code to '.csv') that is to be analyzed must be specified on a newline in the target_list folder. Once again, an example of this is provided in the tests folder. 
 4. If a single output file is used and USE_CUSTOM_SHEETNAME is set to True (the recommended option), the sheetnames for each of the analyzed file must be given in target_list_sheetnames. This is done to allow descriptive file names to be preserved for the log without the sheetnames in excel becoming too long to be unwieldy. The sheetnames are to be put on newlines in the same way as for the target files; the script will simply pull the sheetnames specified in order and apply them to the analyzed data. As before, an example of this is provided in the tests folder.
 
Note that while code work by itself, it also may accept a command-line that points it to the location of a configuration in the form of a yaml file (note, a yaml file is just a attribute-value data type like json or xml).This YAML file contains all of the configuration variables. An example of this is available in the tests folder.
 
## About the \tests directory

This directory will allow a new user to run a functional example of the code, and see an example of how the analyzer is to be configured. If python is installed and added to path; and all of the requirements are installed, I have even provided a bat script that automatically runs it on windows when double-clicked. 

The plan is to use the pytest utility to run a few unit tests on some key functions. It won't be a rigorous job but it's better than nothing, and will hopefully make any future development of this codebase easier. 

Reminder: In order to run all tests, pytest must be executed in the main directory of the project
 
 

## Resources and other notes

The main resource that was used to create the theory for this analyzer is the *Parachute Recovery Systems Design Manual*. The relevant parts of it are summarized in the work term report that I (Artem Sotnikov) have written. 
 
