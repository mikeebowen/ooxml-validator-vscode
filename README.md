[![OOXML Validator VS Code Tests](https://github.com/mikeebowen/ooxml-validator-vscode/actions/workflows/main.yml/badge.svg)](https://github.com/mikeebowen/ooxml-validator-vscode/actions/workflows/main.yml)
[![Coverage Status](https://coveralls.io/repos/github/mikeebowen/ooxml-validator-vscode/badge.svg?branch=main)](https://coveralls.io/github/mikeebowen/ooxml-validator-vscode?branch=main)

# OOXML Validator VSCode Extension

The OOXML Validator validates Open Office XML files to help find errors in XML parts. It is a wrapper around the [Open XML SDK](https://github.com/OfficeDev/Open-XML-SDK)'s validate functionality. It displays the errors found in the xml parts of an Office Open XML file in VSCode and creates an optional CSV or JSON log file of the errors.

## Features

- Validates Office Open XML files against Office 2007, 2010, 2013, 2016, or 2019 _Defaults to 2019_
- Creates optional CSV or JSON log file of the errors.

## Usage

To validate an OOXML file, right click on the file to validate and select "Validate OOXML". This displays the validation errors in the VS Code window. If `ooxml.outPutFilePath` is set, a log of the validation errors is created (defaults to .csv).<sup>[1](#csv-footnote)</sup>

![Demonstration of OOXML Viewer VS Code Extension](/assets/view-errors.gif)

## Extension Settings

This extension contributes the following settings:

- `ooxml.fileFormatVersion`: Number that specifies the version of Office to use to validate OOXML files. Must be in `[2007, 2010, 2013, 2016, 2019]`. Defaults to 2019
- `ooxml.outPutFilePath`:
  String that specifies the absolute filepath to write the output of the validator. If the filenames does not end in .json or .csv, ".csv" will be appended to the filename and saved as a .csv file. Path **MUST** be absolute
- `ooxml.overwriteLogFile`: If true the log file will overwrite previous log files of the same name if they exist. If false a unique timestamp is added to the filename. Default is false.

## Release Notes

### 1.0.0

Initial release of OOXML Validator.

Adds Features

- Validates OOXML files.
- Creates optional .csv or .json file of the errors.

---

<a id="csv-footnote">1</a>: Using [Excel Viewer](https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer) to view the .csv file.
