/* eslint-disable @typescript-eslint/naming-convention */
import { Uri, ViewColumn, WebviewPanel, window, workspace } from 'vscode';
import { spawn } from 'child_process';
import { join, dirname, basename, isAbsolute, normalize, extname } from 'path';
import { createObjectCsvWriter } from 'csv-writer';
import { TextEncoder } from 'util';

interface IValidationError {
  Description?: string
  Path?: {
    NamespacesDefinitions?: string[]
    Namespaces: any
    XPath?: string
    PartUri?: string
  }
  Id?: string
  ErrorType?: number
}

class ValidationError {
  Description?: string;
  NamespacesDefinitions?: string[];
  private _Namespaces: any;
  XPath?: string;
  PartUri?: string;
  Id?: string;
  ErrorType?: number;

  get Namespaces() {
    return JSON.stringify(this._Namespaces).toString();
  }

  constructor(options: IValidationError) {
    this.Id = options.Id;
    this.Description = options.Description;
    this._Namespaces = options.Path?.Namespaces;
    this.NamespacesDefinitions = options.Path?.NamespacesDefinitions;
    this.XPath = options.Path?.XPath;
    this.PartUri = options.Path?.PartUri;
    this.ErrorType = options.ErrorType;
  }
}

interface IHeader {
  id: string
  title: string
}

class Header {
  id: string;
  title: string;
  constructor(options: IHeader) {
    this.id = options.id;
    this.title = options.title;
  }
}

export default class OOXMLValidator {
  static validate = (file: Uri) => {
    const panel: WebviewPanel = window.createWebviewPanel('validateOOXML', 'OOXML Validate', ViewColumn.One, { enableScripts: true });
    panel.webview.html = getWebviewContent();
    const formatVersions: any = {
      '2007': '0',
      '2010': '1',
      '2013': '2',
      '2016': '3',
      '2019': '4',
    };
    const configVersion: number | undefined = workspace.getConfiguration('ooxml').get('fileFormatVersion');
    let json = '';
    let err = '';
    const versionStr = configVersion?.toString();
    const versions = Object.keys(formatVersions);
    // Default to the latest format version
    const version = versionStr && versions.includes(versionStr) ? versionStr : formatVersions[versions[versions.length - 1]];
    const child = spawn(join(__dirname, '..', 'bin', 'OOXMLValidatorCLI.exe'), [file.fsPath, formatVersions[version]]);
    child.stdout.on('data', data => {
      json += data;
    });

    child.stderr.on('data', data => {
      err += data;
    });

    child.on('close', async code => {
      const validationErrors: ValidationError[] = JSON.parse(json).map((j: IValidationError) => new ValidationError(j));
      let content: string;
      try {
        if (err && err.length) {
          await window.showErrorMessage(err, { modal: true });
          panel.dispose();
        } else if (validationErrors.length) {
          const path: string | undefined = workspace.getConfiguration('ooxml').get('outPutFilePath');
          if (path) {
            const normalizedPath = normalize(path);
            if (isAbsolute(normalizedPath)) {
              await workspace.fs.createDirectory(Uri.file(dirname(normalizedPath)));
              const ext = extname(basename(normalizedPath));
              if (ext === '.json') {
                const encoder = new TextEncoder();
                await workspace.fs.writeFile(Uri.file(normalizedPath), encoder.encode(JSON.stringify(validationErrors)));
              } else {
                const fixedPath = ext === '.csv' ? normalizedPath : `${normalizedPath}.csv`;
                const csvWriter = createObjectCsvWriter({
                  path: fixedPath,
                  header: Object.keys(validationErrors[0]).map(e => new Header({ id: e, title: e })),
                });
                await csvWriter.writeRecords(validationErrors);
              }
            } else {
              await window.showErrorMessage('OOXML Validator\nooxml.outPutFilePath must be an absolute path');
            }
          }
          content = getWebviewContent(validationErrors, basename(file.fsPath));
          panel.webview.html = content;
        } else {
          content = getWebviewContent([], basename(file.fsPath));
          panel.webview.html = content;
        }
      } catch (err) {
        await window.showErrorMessage(err.message || err, { modal: true });
        panel.dispose();
      }
    });
  };
}
function getWebviewContent(validationErrors?: ValidationError[], fileName?: string): string {
  if (validationErrors && validationErrors.length) {
    let list = '';
    validationErrors.forEach(err => {
      list += `<dl class="row">
            <dt class="col-sm-3">Id</dt>
            <dd class="col-sm-9">${err.Id}</dd>
            <dt class="col-sm-3">Description</dt>
            <dd class="col-sm-9">${err.Description}</dd>
            <dt class="col-sm-3">XPath</dt>
            <dd class="col-sm-9">
              ${err.XPath}
            </dd>
            <dt class="col-sm-3">Part URI</dt>
            <dd class="col-sm-9">${err.PartUri}</dd>
            <dt class="col-sm-3">NamespacesDefinitions</dt>
            <dd class="col-sm-9">
              <ul>
                ${err.NamespacesDefinitions?.map((n: string) => `<li>${n}</li>`).join('')}
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
          <title>OOXML Validation Errors</title>
          <body>
            <div class="container-fluid pt-3 ol-3">
            <p>
              <button class="btn btn-warn" type="button" data-toggle="collapse" data-target="#collapseExample" aria-expanded="false" aria-controls="collapseExample">
                View Errors
              </button>
            </p>
            <div class="collapse" id="collapseExample">
              <div class="card card-body">
                ${list}
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
              <h1 class="display-4">${fileName} Is Valid!!</h1>
              <p class="lead">OOXML Validator did not find any validation errors in ${fileName}.</p>
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
}
