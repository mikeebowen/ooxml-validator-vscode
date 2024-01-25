[![VS Marketplace](https://vsmarketplacebadges.dev/version-short/mikeebowen.ooxml-validator-vscode.svg)](https://marketplace.visualstudio.com/items?itemName=mikeebowen.ooxml-validator-vscode)
[![Unit Tests](https://github.com/mikeebowen/ooxml-validator-vscode/actions/workflows/main.yml/badge.svg)](https://github.com/mikeebowen/ooxml-validator-vscode/actions/workflows/main.yml)
[![Coverage Status](https://coveralls.io/repos/github/mikeebowen/ooxml-validator-vscode/badge.svg?branch=main)](https://coveralls.io/github/mikeebowen/ooxml-validator-vscode?branch=main)
[![Project License](https://img.shields.io/github/license/mikeebowen/ooxml-validator-vscode?label=license)](https://github.com/mikeebowen/ooxml-validator-vscode/blob/main/LICENSE)

# OOXML Validator VS Code Extension

The OOXML Validator validates Office Open XML files (.docx, .docm, .dotm, .dotx, .pptx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsx, .xlsm, .xltm, .xltx, or .xlam) and displays the validation errors found in the xml parts in VSCode and creates an optional CSV or JSON log file of the validation errors.

## Features

- Validates Office Open XML files against Office Perpetual 2007, 2010, 2013, 2016, 2019, 2021, or Microsoft 365 Subscription _Defaults to Microsoft 365 Subscription_
- Creates optional CSV or JSON log file of the errors.

## Usage

To validate an OOXML file, right click on the file to validate and select "Validate OOXML". This displays the validation errors in the VS Code window. If `ooxml.outPutFilePath` is set, a log of the validation errors is created (defaults to .csv).

![Demonstration of OOXML Viewer VS Code Extension](https://raw.githubusercontent.com/mikeebowen/ooxml-validator-vscode/main/assets/view-errors-2.gif)

<sup>\* Using [Excel Viewer](https://marketplace.visualstudio.com/items?itemName=GrapeCity.gc-excelviewer) to view the .csv file.</sup>

## Extension Settings

This extension contributes the following settings:

- `ooxml.fileFormatVersion`: Number that specifies the version of Office to use to validate OOXML files. Must be in `[2007, 2010, 2013, 2016, 2019, 2021, 365]`. If no value is set or the value is not valid, defaults to 365. For the Office version 365 refers to the subscription product (Microsoft 365), while the other versions refer the perpetual versions for a given year.
- `ooxml.outPutFilePath`:
  String that specifies the absolute file path to write the output of the validator. If the file name does not end in ".json" or ".csv", ".csv" will be appended to the filename and saved as a .csv file. Path **MUST** be absolute. If no value is set, no log file will be saved.
- `ooxml.overwriteLogFile`: If true, the new log file will overwrite previous log file of the same name if it exists. If false, a unique timestamp is added to the filename. Default is false.
- `ooxml.dotNetPath`: The absolute path to the .Net Runtime. Path **MUST** be absolute. If not set, the extension will download the latest dotnet runtime. 


> The path to the dotnet runtime executable will be similar to `"C:\\Program Files\\dotnet\\dotnet.exe"` for Windows and `"/usr/local/share/dotnet/dotnet"` for macos. For more information on finding the path to the dotnet executable, [this Stack Overflow post explains it for Windows](https://stackoverflow.com/questions/42588392/where-is-the-dotnet-command-executable-located-on-windows) and [this post explains how to find it on macos](https://stackoverflow.com/questions/41112054/where-is-the-default-net-core-sdk-path-in-macos).

## Release Notes

Please see the [Changelog](CHANGELOG.md)

---
<p align="right">Logo created with <a href="https://logomakr.com/">LogoMakr.com</a></p>
