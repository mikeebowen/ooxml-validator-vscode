[![Unit Tests](https://github.com/mikeebowen/ooxml-validator-vscode/actions/workflows/main.yml/badge.svg)](https://github.com/mikeebowen/ooxml-validator-vscode/actions/workflows/main.yml)
[![Coverage Status](https://coveralls.io/repos/github/mikeebowen/ooxml-validator-vscode/badge.svg?branch=main)](https://coveralls.io/github/mikeebowen/ooxml-validator-vscode?branch=main)
![GitHub release (latest by date)](https://img.shields.io/github/v/release/mikeebowen/ooxml-validator-vscode)
![GitHub](https://img.shields.io/github/license/mikeebowen/ooxml-validator-vscode?label=license)

# OOXML Validator VSCode Extension

The OOXML Validator validates Office Open XML files (.docx, .xlsx, .pptx) and displays the validation errors found in the xml parts in VSCode and creates an optional CSV or JSON log file of the validation errors.

## Features

- Validates Office Open XML files against Office 2007, 2010, 2013, 2016, or 2019 _Defaults to 2019_
- Creates optional CSV or JSON log file of the errors.

## Requirements

The OOXML Validator requires that the [.NET Install Tool for Extension Authors](https://marketplace.visualstudio.com/items?itemName=ms-dotnettools.vscode-dotnet-runtime) VS Code extension be installed.

## Usage

To validate an OOXML file, right click on the file to validate and select "Validate OOXML". This displays the validation errors in the VS Code window. If `ooxml.outPutFilePath` is set, a log of the validation errors is created (defaults to .csv).

![Demonstration of OOXML Viewer VS Code Extension](https://raw.githubusercontent.com/mikeebowen/ooxml-validator-vscode/main/assets/view-errors.gif)

<sup>\* Using [Excel Viewer](https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer) to view the .csv file.</sup>

## Extension Settings

This extension contributes the following settings:

- `ooxml.fileFormatVersion`: Number that specifies the version of Office to use to validate OOXML files. Must be in `[2007, 2010, 2013, 2016, 2019]`. Defaults to 2019
- `ooxml.outPutFilePath`:
  String that specifies the absolute filepath to write the output of the validator. If the filenames does not end in .json or .csv, ".csv" will be appended to the filename and saved as a .csv file. Path **MUST** be absolute
- `ooxml.overwriteLogFile`: If true the log file will overwrite previous log files of the same name if they exist. If false a unique timestamp is added to the filename. Default is false.

## Release Notes

Please see the [Changelog](CHANGELOG.md)
