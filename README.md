# Config Creator

The goal of this script was to be able to take XLSX data and mass create switch configurations based on the data in the XLSX spreadsheet. I also added in a option for TXT input for one off tasks.

## Getting Started

Look at the example files. In the XLSX spreadsheet, you can add as many columns as you want to the values page, just make sure that the column name is in the variable so it knows to pull data.

Step 1. Build the template. Any data you want to pull from XLSX you can reference with a variable like #VARIABLE1#.
Step 2. Build the XLSX spreadsheet. Do not change the header information in the first 2 pages. Only change header information based on the variable name in the values page.
Step 3. Run and profit

Report any issues to my email and I will get them fixed.

### Prerequisites

GIT (This is required to download the XLHELPER module using a fork that  I made for compatibility with Python 2.7)
XLHELPER
OPENPYXL

## Deployment

Just execute the script and answer the questions

## Features
- Text-based variable replacement
- XLSX-based variable replacement
- Mass produce configurations based on template and data in XLSX spreadsheet

## *Caveats
- None

## Versioning

VERSION 1.1
Currently Implemented Features
-Text based variable replacement
-XLSX based variable replacement
-Automatic math for CIDR on XLSX
	- If the XLSX column data contains a subnet in cidr format, and the variable contains IPADD, it will automatically switch it out for the gateway IP in subnet format
		- e.g. 10.1.1.0/24 -> 10.1.1.1 255.255.255.0
		- Additional variables would be HSRPPRI (changes to 10.1.1.2) and HSRPSEC (changes to 10.1.1.3)
	- Full list of math based variables (IPADD,HSRPPRI,HSRPSEC)

## Authors

* **Matt Cross** - [RouteAllThings](https://github.com/routeallthings)

See also the list of [contributors](https://github.com/routeallthings/Config-Creator/contributors) who participated in this project.

## License

This project is licensed under the GNU - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Thanks to HBS for giving me a reason to write this.
