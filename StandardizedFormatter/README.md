# StandardizedFormatter
This project was worked on in order to solve issues when working between systems that use different formatting within excel files.

The main issue that arose was formatting of street names. Our system needed .xlsx or .csv files with all uppercase and shortened street types (ex. street -> ST, ave -> AVE).

This program uses a regex to match a string specified by a dictionary. Once found, it will replace values within the file with a corresponding value mapped to the initial value within the dictionary.

## User Guide
1. When prompted, enter path to file that should be reformatted.
2. Enter name of address column that should be reformatted.
3. Enter name of save file.
4. Enter "n" or "N" to leave program. (Needs improved functionality)

## Build

The package for this program was built using PyInstaller. The Executable can be found in /dist/

## Notes

file-formatter and file-formatter-1.0 uses a different iterable implementation and is not functional. 
It is only included for personal reference when working with pandas. file-formatter-1.1 should be used.
