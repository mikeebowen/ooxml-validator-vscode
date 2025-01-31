import { spawnSync } from 'child_process';
import { createObjectCsvWriter } from 'csv-writer';
import { basename, dirname, extname, isAbsolute, join, normalize } from 'path';
import { TextEncoder } from 'util';
import { Uri, ViewColumn, WebviewPanel, commands, extensions } from 'vscode';
import { Header, IDotnetAcquireResult, IValidationError, ValidationError } from './models';
import { ExtensionUtilities, WindowUtilities, WorkspaceUtilities } from './utilities';

export default class OOXMLValidator {
  static createLogFile = async (validationErrors: ValidationError[], path: string): Promise<string | undefined> => {
    let normalizedPath = normalize(path);
    let ext = extname(basename(normalizedPath));

    if (ext !== '.csv' && ext !== '.json') {
      normalizedPath = `${normalizedPath}.csv`;
      ext = '.csv';
    }

    if (isAbsolute(normalizedPath)) {
      await WorkspaceUtilities.createDirectory(dirname(normalizedPath));
      const overwriteLogFile: boolean | undefined = WorkspaceUtilities.getConfigurationValue<boolean>('ooxml', 'overwriteLogFile');

      if (!overwriteLogFile) {
        normalizedPath = `${normalizedPath.substring(0, normalizedPath.length - ext.length)}.${new Date()
          .toISOString()
          .replaceAll(':', '_')}${ext}`;
      }

      if (ext === '.json') {
        const encoder = new TextEncoder();
        await WorkspaceUtilities.writeFile(normalizedPath, encoder.encode(JSON.stringify(validationErrors, null, 2)));
      } else {
        const csvWriter = createObjectCsvWriter({
          path: normalizedPath,
          header: Object.keys(validationErrors[0]).map(e => new Header({ id: e, title: e })),
        });

        const errorsForCsv = validationErrors.map((ve: ValidationError) => {
          const copy: { [key: string]: any } = Object.assign({}, ve);

          for (const [key, value] of Object.entries(copy)) {
            const k = key as 'Id' | 'Description' | 'Namespaces' | 'NamespacesDefinitions' | 'XPath' | 'PartUri' | 'ErrorType';
            copy[k] = JSON.stringify(value, null, 2);
          }

          return copy;
        });

        await csvWriter.writeRecords(errorsForCsv);
      }

      return normalizedPath;
    } else {
      await WindowUtilities.showError('OOXML Validator\nooxml.outPutFilePath must be an absolute path');
    }
  };

  static getWebviewContent = (
    validationErrors?: ValidationError[],
    version?: string,
    fileName?: string,
    path?: string,
  ): string => {
    if (validationErrors && validationErrors.length) {
      let list = '';
      validationErrors.forEach((err) => {
        list += `<dl class="row">
              <dt class="col-sm-3">Id</dt>
              <dd class="col-sm-9">${err.Id}</dd>
              <dt class="col-sm-3">Description</dt>
              <dd class="col-sm-9">${err.Description?.replace(
          /</g,
          '&lt;',
        )}</dd>
              <dt class="col-sm-3">XPath</dt>
              <dd class="col-sm-9">
                ${err.XPath}
              </dd>
              <dt class="col-sm-3">Part URI</dt>
              <dd class="col-sm-9">${err.PartUri}</dd>
              <dt class="col-sm-3">NamespacesDefinitions</dt>
              <dd class="col-sm-9">
                <ul>
                  ${err.NamespacesDefinitions?.map(
          (n: string) => `<li>${n}</li>`,
        ).join('')}
                </ul>
              </dd>
            </dl>`;
      });

      return `<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>
            <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
            <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
            <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
            <style>
              label#error-btn:after {
                content:'View Errors';
              }

              [aria-expanded="true"] label#error-btn:after {
                content:'Hide Errors';
              }
            </style>
            <title>OOXML Validation Errors</title>
            <body>
              <div class="container-fluid pt-3 ol-3">
              <div class="row pb-3">
                <div class="col">
                  <h1>There ${validationErrors.length === 1 ? 'was' : 'were'} ${
        validationErrors.length
      } Validation Error${validationErrors.length > 1 ? 's' : ''} Found</h1>
                  <h2>Validating against ${version}</h2>
  ${
    path
      ? `<h3>A log of these errors was saved as "${path}"</h3>`
      : // eslint-disable-next-line max-len
      '<h3>No log of these errors was saved.</h3><h4>Set "ooxml.outPutFilePath" in settings.json to save a log (csv or json) of the errors</h4>'
  }
                </div>
              </div>
              <div class="row pb-3">
                <div class="col">
                  <div>
                    <div class="card card-body">
                      ${list}
                    </div>
                  </div>
                </div>
              </div>
            </body>
        </html>`;
    } else if (validationErrors && !validationErrors.length) {
      return `<!DOCTYPE html>
      <html lang="en">
      <head>
          <meta charset="UTF-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
          <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>
          <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>
          <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>
          <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>
          <title>OOXML Validation Errors</title>
          <body>
            <div class="container pt-3 ol-3">
              <div class="row">
                <div class="col">
                <div class="jumbotron">
                <h1 class="display-4 text-center">No Errors Found!!</h1>
                <h2>Validating Against ${version}</h2>
                <p class="lead text-center">OOXML Validator did not find any validation errors in ${fileName}.</p>
              </div>
                </div>
              </div>
            </div>
          </body>
        </html>`;
    }

    return `<!DOCTYPE html>
            <html lang="en">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <style>
                .container {
                  display: -webkit-box;
                  display: -ms-flexbox;
                  display: flex;
                  -webkit-box-align: center;
                      -ms-flex-align: center;
                          align-items: center;
                  -webkit-box-pack: center;
                      -ms-flex-pack: center;
                          justify-content: center;
                  min-height: 100vh;
                  background-color: #ededed;
                }

                .loader {
                  max-width: 15rem;
                  width: 100%;
                  height: auto;
                  stroke-linecap: round;
                }

                circle {
                  fill: none;
                  stroke-width: 3.5;
                  -webkit-animation-name: preloader;
                          animation-name: preloader;
                  -webkit-animation-duration: 3s;
                          animation-duration: 3s;
                  -webkit-animation-iteration-count: infinite;
                          animation-iteration-count: infinite;
                  -webkit-animation-timing-function: ease-in-out;
                          animation-timing-function: ease-in-out;
                  -webkit-transform-origin: 170px 170px;
                          transform-origin: 170px 170px;
                  will-change: transform;
                }
                circle:nth-of-type(1) {
                  stroke-dasharray: 550;
                }
                circle:nth-of-type(2) {
                  stroke-dasharray: 500;
                }
                circle:nth-of-type(3) {
                  stroke-dasharray: 450;
                }
                circle:nth-of-type(4) {
                  stroke-dasharray: 300;
                }
                circle:nth-of-type(1) {
                  -webkit-animation-delay: -0.15s;
                          animation-delay: -0.15s;
                }
                circle:nth-of-type(2) {
                  -webkit-animation-delay: -0.3s;
                          animation-delay: -0.3s;
                }
                circle:nth-of-type(3) {
                  -webkit-animation-delay: -0.45s;
                  -moz-animation-delay:  -0.45s;
                          animation-delay: -0.45s;
                }
                circle:nth-of-type(4) {
                  -webkit-animation-delay: -0.6s;
                  -moz-animation-delay: -0.6s;
                          animation-delay: -0.6s;
                }

                @-webkit-keyframes preloader {
                  50% {
                    -webkit-transform: rotate(360deg);
                            transform: rotate(360deg);
                  }
                }

                @keyframes preloader {
                  50% {
                    -webkit-transform: rotate(360deg);
                            transform: rotate(360deg);
                  }
                }
                </style>
                <title>OOXML Validation Errors</title>
            </head>
            <body>
            <div class="container">
            <div style="display: block">
              <div style="display: block">
                 <h1>Validating OOXML File</h1>
              </div>
                  <svg class="loader" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 340 340">
                    <circle cx="170" cy="170" r="160" stroke="#E2007C"/>
                    <circle cx="170" cy="170" r="135" stroke="#404041"/>
                    <circle cx="170" cy="170" r="110" stroke="#E2007C"/>
                    <circle cx="170" cy="170" r="85" stroke="#404041"/>
                  </svg>
                </div>

          </div>
            </body>
            </html>`;
  };

  static validate = async (uri: Uri) => {
    const panel: WebviewPanel = WindowUtilities.createWebView('validateOOXML', 'OOXML Validate', ViewColumn.One, { enableScripts: true });

    try {
      panel.webview.html = OOXMLValidator.getWebviewContent();

      const formatVersions: Map<string, string> = new Map();
      formatVersions.set('2007', 'Office2007');
      formatVersions.set('2010', 'Office2010');
      formatVersions.set('2013', 'Office2013');
      formatVersions.set('2016', 'Office2016');
      formatVersions.set('2019', 'Office2019');
      formatVersions.set('2021', 'Office2021');
      formatVersions.set('365', 'Microsoft365');

      const configVersion: number | undefined = WorkspaceUtilities.getConfigurationValue<number>('ooxml', 'fileFormatVersion');
      const versionStr = configVersion?.toString();
      const versions = [...formatVersions.keys()];
      // Default to the latest format version
      const version =
        versionStr && versions.includes(versionStr) ? formatVersions.get(versionStr) : [...formatVersions.values()].pop();

      await commands.executeCommand('dotnet.showAcquisitionLog');

      const requestingExtensionId = 'mikeebowen.ooxml-validator-vscode';
      let dotnetPath: string | undefined = WorkspaceUtilities.getConfigurationValue<string>('ooxml', 'dotNetPath');

      if (dotnetPath) {
        const dotnetPathIsValid = await ExtensionUtilities.isDotNetRuntime(dotnetPath.toString());

        if (!dotnetPathIsValid) {
          dotnetPath = undefined;
          await WindowUtilities.showWarning(
            'OOXML Validator: The .NET path set in the settings is not valid.',
            // eslint-disable-next-line max-len
            'Using the .NET Install Tool for Extension Authors extension to acquire .NET Runtime.\nUpdate the ooxml.dotNetPath setting to the use a local runtime.',
            true,
          );
        }
      }

      if (!dotnetPath) {
        const commandRes = await commands.executeCommand<IDotnetAcquireResult>('dotnet.acquire', {
          version: '8.0',
          requestingExtensionId,
        });
        dotnetPath = commandRes!.dotnetPath;

        if (!dotnetPath) {
          throw new Error('Could not acquire .NET runtime.');
        }
      }

      const ooxmlValidateExtension = extensions.getExtension(requestingExtensionId);

      if (!ooxmlValidateExtension) {
        throw new Error('Could not find OOXML Validate extension.');
      }

      const ooxmlValidateLocation = join(ooxmlValidateExtension.extensionPath, 'OOXMLValidator', 'OOXMLValidatorCLI.dll');
      const ooxmlValidateArgs = [ooxmlValidateLocation, uri.fsPath, version];

      // This will install any missing Linux dependencies.
      await commands.executeCommand('dotnet.ensureDotnetDependencies', { command: dotnetPath, arguments: ooxmlValidateArgs });

      const result = spawnSync(dotnetPath, ooxmlValidateArgs as string[]);
      const stderr = result?.stderr?.toString();

      if (stderr?.length > 0) {
        WindowUtilities.showError(`Failed to run OOXML Validator. The error was:\n${stderr}`, true);
        panel.dispose();

        return;
      }

      const appOutput = result?.stdout?.toString();
      const validationErrors: ValidationError[] = JSON.parse(appOutput).map((r: IValidationError) => new ValidationError(r));
      let content: string;

      if (validationErrors.length) {
        const path: string | undefined = WorkspaceUtilities.getConfigurationValue<string>('ooxml', 'outPutFilePath');
        let pathToSavedFile: string | undefined;

        if (path) {
          pathToSavedFile = await OOXMLValidator.createLogFile(validationErrors, path);
        }

        content = OOXMLValidator.getWebviewContent(validationErrors, version, basename(uri.fsPath), pathToSavedFile);
        panel.webview.html = content;
      } else {
        content = OOXMLValidator.getWebviewContent([], version, basename(uri.fsPath));
        panel.webview.html = content;
      }
    } catch (error: any) {
      let errMsg = error.message || error;

      Object.values(error).some((e: any) => {
        const str = e.toString && e.toString();

        if (str && str.includes('dotnet.')) {
          errMsg =
            // eslint-disable-next-line max-len
            'The ".NET Install Tool for Extension Authors" VS Code extensionMUST be installed\nor the ooxml.dotNetPath must be set to th absolute path to the .Net Runtime\nfor the OOXML Validator extension to work.';

          return true;
        }
      });

      panel?.dispose();

      await WindowUtilities.showError(errMsg, true);
    }
  };
}
