# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [1.4.5] - 2022-04-13

### Fixed
- Updated minimist to version 1.2.6

## [1.4.4] - 2022-04-13

## Added
- Added LogoMakr.com link to README.md

## [1.4.3] - 2022-01-24

## Added
- Updated mocha to fix security vulnerability
- Moved classes and interfaces into separate files

## [1.4.2] - 2022-01-20

### Fixed
- Fixed versioning issue

## [1.4.1] - 2022-01-20

### Added
- Updated README

## [1.4.0] - 2022-01-20

### Added
- Validation for .docm, .dotm, .dotx, .pptm, .potm, .potx, .ppam, .ppsm, .ppsx, .xlsm, .xltm, .xltx, and .xlam files, closes [#8](https://github.com/mikeebowen/ooxml-validator-vscode/issues/8)
- CLI errors show in modal window

### Fixed
- Web Panel closes if there is an error from the CLI

## [1.3.1] - 2022-01-06

### Fixed
- Updated dependencies with security vulnerabilities

## [1.3.0] - 2022-01-05

### Added
- Support for validation against Microsoft 365 subscription

### Fixed
- Validation no longer always picks 2021

## [1.2.0] - 2021-12-01

### Added

- Ability to validate against Office 2021.
- The Office version being validated against is displayed in the web view.

## [1.1.0] - 2021-11-30

### Added

- Configuration option to manually set path to .Net Runtime, closes [#7](https://github.com/mikeebowen/ooxml-validator-vscode/issues/7).

## [1.0.2] - 2021-09-29

### Fixed

- List of possible elements expected is omitted [#4](https://github.com/mikeebowen/ooxml-validator-vscode/issues/4).

### Added

- Updated README with explanation of how to skip creating the log file.

## [1.0.1] - 2021-09-03

### Fixed

- Display correct error message when dotnet runtime extension is missing [#2](https://github.com/mikeebowen/ooxml-validator-vscode/issues/2).

### Added

- Updated README badges
- Created my free logo at [LogoMakr.com](https://logomakr.com/).

## [1.0.0] - 2021-09-02

### Added

- Validates OOXML files.
- Creates optional .csv or .json file of the errors.
