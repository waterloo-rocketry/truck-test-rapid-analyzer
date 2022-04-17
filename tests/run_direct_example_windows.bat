:: To be run only from the file directory in which it is located, otherwise the cds command 
:: will not function correctly and the bat file will not work

@echo off
echo We need to go up one directory for this to work correctly 
@echo on
cd .. 
python -m tests.rapid_analyzer_example_direct 
@echo off
echo:
echo a 'cmd /k' command is run simply to keep the command shell open
@echo on
cmd /k