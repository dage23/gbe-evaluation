# Python Exercise
Program made in python to monitor a desired folder, using watchdog, when a new file is detected, it verifies if the file is .xls. If the file is .xls, 
it consolidates the new file (copies each of the sheets of the new file to the master file) with a master file created in the /Processed/ folder. 
If not, it is moved to the Not Applicable folders.

The Processed and Not Applicable folders are created in the folder that is being monitored.

The Executable file is stored in the /dist folder.
