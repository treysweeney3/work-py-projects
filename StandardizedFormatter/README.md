# StandardizedFormatter
This project was worked on in order to solve issues when working between systems that use different formatting within excel files.

The main issue that arose was formatting of street names. Our system needed .xlsx or .csv files with all uppercase and shortened street types (ex. street -> ST, ave -> AVE).

This program uses a regex to match a string specified by a dictionary. Once found, it will replace values within the file with a corresponding value mapped to the initial value within the dictionary.

More explanation on functionality will be added at a later date.
