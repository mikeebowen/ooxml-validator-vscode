/* eslint-disable @typescript-eslint/naming-convention */
import { commands, Extension, extensions, Uri, ViewColumn, WebviewPanel, window, workspace, WorkspaceConfiguration } from 'vscode';
import { SinonStub, spy, stub } from 'sinon';
import { expect, use } from 'chai';
import * as shallowDeepEqual from 'chai-shallow-deep-equal';
import * as path from 'path';
import * as child_process from 'child_process';
import OOXMLValidator, { ValidationError, effEss, IValidationError } from '../../ooxml-validator';
import { basename, normalize, join } from 'path';
import { TextEncoder } from 'util';
import * as csvWriter from 'csv-writer';
import { isEqual } from 'lodash';

use(shallowDeepEqual);

suite('OOXMLValidator', function () {
  this.timeout(15000);
  const stubs: SinonStub[] = [];

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  suite('createLogFile', function () {
    test('should throw an error if path is not absolute', async function () {
      const isAbsoluteStub = stub(path, 'isAbsolute').returns(false);
      const showErrorMessageStub = stub(window, 'showErrorMessage').returns(Promise.resolve() as Thenable<undefined>);
      stubs.push(isAbsoluteStub, showErrorMessageStub);
      await OOXMLValidator.createLogFile([], 'tacocat');
      expect(isAbsoluteStub.args[0][0]).to.eq('tacocat.csv');
      expect(showErrorMessageStub.calledOnce).to.be.true;
    });

    test('should create a json file if the file\'s extension is ".json"', function (done) {
      const testPath: string = normalize(path.join(__dirname, 'racecar.json'));
      const createDirectoryStub = stub(effEss, 'createDirectory');
      const writeFileStub: SinonStub = stub(effEss, 'writeFile');
      stubs.push(writeFileStub, createDirectoryStub);
      const errors: ValidationError[] = [new ValidationError({}), new ValidationError({})];
      const testBuffer = new TextEncoder().encode(JSON.stringify(errors, null, 2));

      OOXMLValidator.createLogFile(errors, testPath)
        .then(() => {
          expect(createDirectoryStub.calledOnceWith(Uri.file(path.dirname(testPath)))).to.be.true;
          expect(writeFileStub.args[0][0].path.includes(normalize(path.join(__dirname, 'racecar'))));
          expect(Buffer.from(testBuffer).equals(Buffer.from(writeFileStub.args[0][1]))).to.be.true;
          done();
        })
        .catch(err => {
          done(err);
        });
    });

    test('should create a csv file if json is not specified', async function () {
      const testPath: string = normalize(path.join(__dirname, 'racecar.foo'));
      const errors: ValidationError[] = [new ValidationError({}), new ValidationError({})];
      const writeRecordsSpy = spy();
      const createObjectCsvWriterStub = stub(csvWriter, 'createObjectCsvWriter').callsFake((params: any): any => {
        expect(params?.path.endsWith('.csv')).to.eq(true, 'params?.path.endsWith');
        return {
          writeRecords: writeRecordsSpy,
        };
      });

      stubs.push(createObjectCsvWriterStub);
      await OOXMLValidator.createLogFile(errors, testPath);
      expect(createObjectCsvWriterStub.calledOnce).to.eq(true, 'createObjectCsvWriterStub.calledOnce');
      expect(writeRecordsSpy.calledOnce).to.eq(true, 'writeRecordsSpy.calledOnce');
      expect(writeRecordsSpy.args[0][0].every((t: any, i: number) => isEqual(t, Object.assign({}, errors[i])))).to.eq(
        true,
        'writeRecordsSpy.args.every',
      );
    });
  });

  suite('getWebViewContent', function () {
    test('should return the errors in html if there are validation errors', function (done) {
      const validationErrors = [
        new ValidationError({
          Id: '1',
          Description: 'the first test error',
          Path: {
            NamespacesDefinitions: ['firstNamespace', 'secondNamespace'],
            Namespaces: {},
            PartUri: 'some/uri',
          },
        }),
        new ValidationError({
          Id: '2',
          Description: 'the second test error',
          Path: {
            NamespacesDefinitions: ['thirdNamespace', 'fourthNamespace'],
            Namespaces: {},
            PartUri: 'some/other/uri',
          },
        }),
      ];
      const testHtml =
        '<!DOCTYPE html>\n        <html lang="en">\n        <head>\n            <meta charset="UTF-8">\n            <meta name="viewport" content="width=device-width, initial-scale=1.0">\n            <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>\n            <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>\n            <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>\n            <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>\n            <style>\n              label#error-btn:after {\n                content:\'View Errors\';\n              }\n\n              [aria-expanded="true"] label#error-btn:after {\n                content:\'Hide Errors\';\n              }\n            </style>\n            <title>OOXML Validation Errors</title>\n            <body>\n              <div class="container-fluid pt-3 ol-3">\n              <div class="row pb-3">\n                <div class="col">\n                  <h1>There Were 2 Validation Errors Found</h1>\n                  <h2>Validating against test-file.csv</h2>\n  <h3>No log of these errors was saved.</h3><h4>Set "ooxml.outPutFilePath" in settings.json to save a log (csv or json) of the errors</h4>\n                </div>\n              </div>\n              <div class="row pb-3">\n                <div class="col">\n                <div class="btn-group-toggle"\n                  data-toggle="collapse"\n                  data-target="#collapseExample"\n                  aria-expanded="false"\n                  aria-controls="collapseExample"\n                >\n                  <label class="btn btn-outline-secondary" id="error-btn">\n                    <input\n                      class="btn btn-outline-secondary"\n                      type="checkbox"\n                      checked\n                    />\n                  </label>\n                  </div>\n                </div>\n              </div>\n              <div class="row pb-3">\n                <div class="col">\n                  <div class="collapse" id="collapseExample">\n                    <div class="card card-body">\n                      <dl class="row">\n              <dt class="col-sm-3">Id</dt>\n              <dd class="col-sm-9">1</dd>\n              <dt class="col-sm-3">Description</dt>\n              <dd class="col-sm-9">the first test error</dd>\n              <dt class="col-sm-3">XPath</dt>\n              <dd class="col-sm-9">\n                undefined\n              </dd>\n              <dt class="col-sm-3">Part URI</dt>\n              <dd class="col-sm-9">some/uri</dd>\n              <dt class="col-sm-3">NamespacesDefinitions</dt>\n              <dd class="col-sm-9">\n                <ul>\n                  <li>firstNamespace</li><li>secondNamespace</li>\n                </ul>\n              </dd>\n            </dl><dl class="row">\n              <dt class="col-sm-3">Id</dt>\n              <dd class="col-sm-9">2</dd>\n              <dt class="col-sm-3">Description</dt>\n              <dd class="col-sm-9">the second test error</dd>\n              <dt class="col-sm-3">XPath</dt>\n              <dd class="col-sm-9">\n                undefined\n              </dd>\n              <dt class="col-sm-3">Part URI</dt>\n              <dd class="col-sm-9">some/other/uri</dd>\n              <dt class="col-sm-3">NamespacesDefinitions</dt>\n              <dd class="col-sm-9">\n                <ul>\n                  <li>thirdNamespace</li><li>fourthNamespace</li>\n                </ul>\n              </dd>\n            </dl>\n                    </div>\n                  </div>\n                </div>\n              </div>\n            </body>\n        </html>';
      const html = OOXMLValidator.getWebviewContent(validationErrors, 'test-file.csv', '/test-directory');
      expect(testHtml).to.equal(html);
      done();
    });

    test('should return the success html if there are no validation errors', function (done) {
      const validationErrors: ValidationError[] = [];
      const testHtml =
        '<!DOCTYPE html>\n      <html lang="en">\n      <head>\n          <meta charset="UTF-8">\n          <meta name="viewport" content="width=device-width, initial-scale=1.0">\n          <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>\n          <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>\n          <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>\n          <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>\n          <title>OOXML Validation Errors</title>\n          <body>\n            <div class="container pt-3 ol-3">\n              <div class="row">\n                <div class="col">\n                <div class="jumbotron">\n                <h1 class="display-4 text-center">No Errors Found!!</h1>\n                <h2>Validating Against test-file.csv</h2>\n                <p class="lead text-center">OOXML Validator did not find any validation errors in /test-directory.</p>\n              </div>\n                </div>\n              </div>\n            </div>\n          </body>\n        </html>';
      const html = OOXMLValidator.getWebviewContent(validationErrors, 'test-file.csv', '/test-directory');
      expect(testHtml).to.equal(html);
      done();
    });

    test('should return the loading html if the validation errors parameter is undefined ', function (done) {
      const testHtml =
        '<!DOCTYPE html>\n            <html lang="en">\n            <head>\n                <meta charset="UTF-8">\n                <meta name="viewport" content="width=device-width, initial-scale=1.0">\n                <style>\n                .container {\n                  display: -webkit-box;\n                  display: -ms-flexbox;\n                  display: flex;\n                  -webkit-box-align: center;\n                      -ms-flex-align: center;\n                          align-items: center;\n                  -webkit-box-pack: center;\n                      -ms-flex-pack: center;\n                          justify-content: center;\n                  min-height: 100vh;\n                  background-color: #ededed;\n                }\n\n                .loader {\n                  max-width: 15rem;\n                  width: 100%;\n                  height: auto;\n                  stroke-linecap: round;\n                }\n\n                circle {\n                  fill: none;\n                  stroke-width: 3.5;\n                  -webkit-animation-name: preloader;\n                          animation-name: preloader;\n                  -webkit-animation-duration: 3s;\n                          animation-duration: 3s;\n                  -webkit-animation-iteration-count: infinite;\n                          animation-iteration-count: infinite;\n                  -webkit-animation-timing-function: ease-in-out;\n                          animation-timing-function: ease-in-out;\n                  -webkit-transform-origin: 170px 170px;\n                          transform-origin: 170px 170px;\n                  will-change: transform;\n                }\n                circle:nth-of-type(1) {\n                  stroke-dasharray: 550;\n                }\n                circle:nth-of-type(2) {\n                  stroke-dasharray: 500;\n                }\n                circle:nth-of-type(3) {\n                  stroke-dasharray: 450;\n                }\n                circle:nth-of-type(4) {\n                  stroke-dasharray: 300;\n                }\n                circle:nth-of-type(1) {\n                  -webkit-animation-delay: -0.15s;\n                          animation-delay: -0.15s;\n                }\n                circle:nth-of-type(2) {\n                  -webkit-animation-delay: -0.3s;\n                          animation-delay: -0.3s;\n                }\n                circle:nth-of-type(3) {\n                  -webkit-animation-delay: -0.45s;\n                  -moz-animation-delay:  -0.45s;\n                          animation-delay: -0.45s;\n                }\n                circle:nth-of-type(4) {\n                  -webkit-animation-delay: -0.6s;\n                  -moz-animation-delay: -0.6s;\n                          animation-delay: -0.6s;\n                }\n\n                @-webkit-keyframes preloader {\n                  50% {\n                    -webkit-transform: rotate(360deg);\n                            transform: rotate(360deg);\n                  }\n                }\n\n                @keyframes preloader {\n                  50% {\n                    -webkit-transform: rotate(360deg);\n                            transform: rotate(360deg);\n                  }\n                }\n                </style>\n                <title>OOXML Validation Errors</title>\n            </head>\n            <body>\n            <div class="container">\n            <div style="display: block">\n              <div style="display: block">\n                 <h1>Validating OOXML File</h1>\n              </div>\n                  <svg class="loader" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 340 340">\n                    <circle cx="170" cy="170" r="160" stroke="#E2007C"/>\n                    <circle cx="170" cy="170" r="135" stroke="#404041"/>\n                    <circle cx="170" cy="170" r="110" stroke="#E2007C"/>\n                    <circle cx="170" cy="170" r="85" stroke="#404041"/>\n                  </svg>\n                </div>\n\n          </div>\n            </body>\n            </html>';
      const html = OOXMLValidator.getWebviewContent();
      expect(testHtml).to.equal(html);
      done();
    });
  });

  suite('validate', function () {
    test('should show validation errors in the web view and use a different version of Office if one is assigned', async function () {
      const sdkValidationErrors = [
        {
          Description: "The 'uri' attribute is not declared.",
          Path: {
            NamespacesDefinitions: ['xmlns:c=\\"http://schemas.openxmlformats.org/drawingml/2006/chart\\"'],
            Namespaces: {},
            XPath: '/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]',
            PartUri: '/word/charts/chart3.xml',
          },
          Id: 'Sch_UndeclaredAttribute',
          ErrorType: 0,
        },
        {
          Description:
            "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://sch…entExpectingComplex",
          ErrorType: 0,
        },
        {
          Description: "The element has unexpected child element 'http://schemas.microsoft.com/office/drawing/2012/chart:leaderLines'.",
          Path: {
            NamespacesDefinitions: ['xmlns:c=\\"http://schemas.openxmlformats.org/drawingml/2006/chart\\"'],
            Namespaces: {},
            XPath: '/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:scatterChart[1]/c:ser[1]/c:dLbls[1]/c:extLst[1]/c:ext[1]',
            PartUri: '/word/charts/chart1.xml',
          },
          Id: 'Sch_UnexpectedElementContentExpectingComplex',
          ErrorType: 0,
        },
      ];
      const validationErrors = sdkValidationErrors.map((v: IValidationError) => new ValidationError(v));
      const testHtml = '<span>hello world</span>';
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns(testHtml);
      const disposeSpy = spy();
      const webview = { html: '' };
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview,
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
        get(key: string) {
          switch (key) {
            case 'fileFormatVersion':
              return 2010;
              break;
            case 'outPutFilePath':
              return undefined;
            default:
              break;
          }
        },
      } as unknown as WorkspaceConfiguration);
      const dotnetPath = 'road to nowhere';
      const extensionPath = 'foobar';

      const executeCommandStub = stub(commands, 'executeCommand').returns(Promise.resolve({ dotnetPath }));

      const getExtensionStub = stub(extensions, 'getExtension').returns({ extensionPath } as Extension<unknown>);

      const spawnSyncStub = stub(child_process, 'spawnSync').returns({
        stdout: Buffer.from(JSON.stringify(sdkValidationErrors)),
        stderr: Buffer.from(''),
        pid: 7,
        output: [null],
        status: 13,
        signal: null,
      });

      stubs.push(
        showErrorMessageStub,
        getWebviewContentStub,
        createWebviewPanelStub,
        getConfigurationStub,
        executeCommandStub,
        getExtensionStub,
        spawnSyncStub,
      );

      const file = Uri.file(__filename);
      const requestingExtensionId = 'mikeebowen.ooxml-validator-vscode';
      const executeCommandArgs = [join('foobar', 'OOXMLValidator', 'OOXMLValidatorCLI.dll')];

      await OOXMLValidator.validate(file);

      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(false, 'panel.dispose() should not have been called');
      expect(showErrorMessageStub.callCount).to.eq(0);
      expect(getWebviewContentStub.getCall(1).firstArg).to.deep.eq(validationErrors);
      expect(getWebviewContentStub.getCall(1).args[1]).to.eq('Office2010');
      expect(getWebviewContentStub.getCall(1).args[2]).to.eq(basename(file.fsPath));
      expect(getWebviewContentStub.getCall(1).args[3]).to.be.undefined;
      expect(webview.html).to.eq(testHtml);
      expect(getExtensionStub.getCall(0).firstArg).to.eq(requestingExtensionId);
      expect(executeCommandStub.getCall(0).firstArg).to.eq('dotnet.showAcquisitionLog');
      expect(executeCommandStub.getCall(1).firstArg).to.eq('dotnet.acquire');
      expect(executeCommandStub.getCall(2).lastArg).to.shallowDeepEqual({
        command: dotnetPath,
        arguments: executeCommandArgs,
      });
      expect(spawnSyncStub.firstCall.args[0]).to.eq(dotnetPath);
      expect(spawnSyncStub.firstCall.args[1]).to.shallowDeepEqual(executeCommandArgs);
    });

    // eslint-disable-next-line max-len
    test('should show validation errors in the web view, create a file with the contents and use a different version of Office if one is assigned', async function () {
      const sdkValidationErrors = [
        {
          Description: "The 'uri' attribute is not declared.",
          Path: {
            NamespacesDefinitions: ['xmlns:c=\\"http://schemas.openxmlformats.org/drawingml/2006/chart\\"'],
            Namespaces: {},
            XPath: '/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]',
            PartUri: '/word/charts/chart3.xml',
          },
          Id: 'Sch_UndeclaredAttribute',
          ErrorType: 0,
        },
        {
          Description:
            "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://sch…entExpectingComplex",
          ErrorType: 0,
        },
        {
          Description: "The element has unexpected child element 'http://schemas.microsoft.com/office/drawing/2012/chart:leaderLines'.",
          Path: {
            NamespacesDefinitions: ['xmlns:c=\\"http://schemas.openxmlformats.org/drawingml/2006/chart\\"'],
            Namespaces: {},
            XPath: '/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:scatterChart[1]/c:ser[1]/c:dLbls[1]/c:extLst[1]/c:ext[1]',
            PartUri: '/word/charts/chart1.xml',
          },
          Id: 'Sch_UnexpectedElementContentExpectingComplex',
          ErrorType: 0,
        },
      ];
      const validationErrors = sdkValidationErrors.map((v: any) => new ValidationError(v));
      const testHtml = '<span>hello world</span>';
      const testFilePath = 'C:\\source\\test\\errors\\tacocat.csv';
      const logFilePath = 'a/logfile/path';
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns(testHtml);
      const disposeSpy = spy();
      const webview = { html: '' };
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview,
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
        get(key: string) {
          switch (key) {
            case 'fileFormatVersion':
              return '2013';
              break;
            case 'outPutFilePath':
              return testFilePath;
            default:
              break;
          }
        },
      } as unknown as WorkspaceConfiguration);
      const file = Uri.file(__filename);
      const createLogFileStub = stub(OOXMLValidator, 'createLogFile').returns(Promise.resolve(logFilePath));

      const extensionPath = 'foobar';
      const getExtensionStub = stub(extensions, 'getExtension').returns({ extensionPath } as Extension<unknown>);

      const requestingExtensionId = 'mikeebowen.ooxml-validator-vscode';
      const dotnetPath = 'road to nowhere';
      const executeCommandStub = stub(commands, 'executeCommand').returns(Promise.resolve({ dotnetPath }));

      const spawnSyncStub = stub(child_process, 'spawnSync').returns({
        stdout: Buffer.from(JSON.stringify(sdkValidationErrors)),
        stderr: Buffer.from(''),
        pid: 7,
        output: [null],
        status: 13,
        signal: null,
      });

      const commandArgs = [join('foobar', 'OOXMLValidator', 'OOXMLValidatorCLI.dll')];

      stubs.push(
        showErrorMessageStub,
        getWebviewContentStub,
        createWebviewPanelStub,
        getConfigurationStub,
        createLogFileStub,
        getExtensionStub,
        spawnSyncStub,
        executeCommandStub,
      );

      await OOXMLValidator.validate(file);

      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(false, 'panel.dispose() should not have been called');
      expect(getWebviewContentStub.getCall(1).firstArg).to.deep.eq(validationErrors);
      expect(getWebviewContentStub.getCall(1).args[1]).to.eq('Office2013');
      expect(getWebviewContentStub.getCall(1).args[2]).to.eq(basename(file.fsPath));
      expect(getWebviewContentStub.getCall(1).args[3]).to.eq(logFilePath);
      expect(createLogFileStub.firstCall.firstArg).to.deep.eq(validationErrors);
      expect(createLogFileStub.firstCall.args[1]).to.eq(testFilePath);
      expect(webview.html).to.eq(testHtml);
      expect(getExtensionStub.getCall(0).firstArg).to.eq(requestingExtensionId);
      expect(executeCommandStub.getCall(0).firstArg).to.eq('dotnet.showAcquisitionLog');
      expect(executeCommandStub.getCall(1).firstArg).to.eq('dotnet.acquire');
      expect(executeCommandStub.getCall(2).lastArg).to.shallowDeepEqual({
        command: dotnetPath,
        arguments: commandArgs,
      });
      expect(spawnSyncStub.firstCall.args[0]).to.eq(dotnetPath);
      expect(spawnSyncStub.firstCall.args[1]).to.shallowDeepEqual(commandArgs);
    });

    test('should show the no errors view if there are no validation errors', async function () {
      const validationErrors: ValidationError[] = [];
      const testHtml = '<span>hello world</span>';
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns(testHtml);
      const disposeSpy = spy();
      const webview = { html: '' };
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview,
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
        get(key: string) {
          switch (key) {
            case 'fileFormatVersion':
              return '2019';
              break;
            case 'outPutFilePath':
              return undefined;
            default:
              break;
          }
        },
      } as unknown as WorkspaceConfiguration);
      const file = Uri.file(__filename);

      const extensionPath = 'foobar';
      const getExtensionStub = stub(extensions, 'getExtension').returns({ extensionPath } as Extension<unknown>);

      const requestingExtensionId = 'mikeebowen.ooxml-validator-vscode';
      const dotnetPath = 'road to nowhere';
      const executeCommandStub = stub(commands, 'executeCommand').returns(Promise.resolve({ dotnetPath }));

      const spawnSyncStub = stub(child_process, 'spawnSync').returns({
        stdout: Buffer.from(JSON.stringify([])),
        stderr: Buffer.from(''),
        pid: 7,
        output: [null],
        status: 13,
        signal: null,
      });
      const commandArgs = [join('foobar', 'OOXMLValidator', 'OOXMLValidatorCLI.dll')];

      stubs.push(
        showErrorMessageStub,
        getWebviewContentStub,
        createWebviewPanelStub,
        getConfigurationStub,
        getExtensionStub,
        executeCommandStub,
        spawnSyncStub,
      );

      await OOXMLValidator.validate(file);

      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(false, 'panel.dispose() should not have been called');
      expect(showErrorMessageStub.callCount).to.eq(0);
      expect(getWebviewContentStub.getCall(1).firstArg).to.deep.eq(validationErrors);
      expect(getWebviewContentStub.getCall(1).args[1]).to.eq('Office2019');
      expect(getWebviewContentStub.getCall(1).args[2]).to.eq(basename(file.fsPath));
      expect(getWebviewContentStub.getCall(1).args[3]).to.be.undefined;
      expect(webview.html).to.eq(testHtml);
      expect(getExtensionStub.getCall(0).firstArg).to.eq(requestingExtensionId);
      expect(executeCommandStub.getCall(0).firstArg).to.eq('dotnet.showAcquisitionLog');
      expect(executeCommandStub.getCall(1).firstArg).to.eq('dotnet.acquire');
      expect(executeCommandStub.getCall(2).lastArg).to.shallowDeepEqual({
        command: dotnetPath,
        arguments: commandArgs,
      });
      expect(spawnSyncStub.firstCall.args[0]).to.eq(dotnetPath);
      expect(spawnSyncStub.firstCall.args[1]).to.shallowDeepEqual(commandArgs);
    });

    test('should throw an error if dotnetPath is undefined', async function () {
      const executeCommandStub = stub(commands, 'executeCommand').returns(Promise.resolve({}));
      const file = Uri.file(__filename);
      const testHtml = '<span>hello world</span>';
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns(testHtml);
      const disposeSpy = spy();
      const webview = { html: '' };
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview,
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
        get(key: string) {
          switch (key) {
            case 'fileFormatVersion':
              return '2010';
              break;
            case 'outPutFilePath':
              return undefined;
            default:
              break;
          }
        },
      } as unknown as WorkspaceConfiguration);

      stubs.push(executeCommandStub, showErrorMessageStub, getWebviewContentStub, createWebviewPanelStub, getConfigurationStub);

      await OOXMLValidator.validate(file);

      expect(disposeSpy.called).to.eq(true, 'panel.dispose() should have been called');
      expect(showErrorMessageStub.firstCall.firstArg).to.eq('Could not resolve the dotnet path!');
    });

    test('should throw an error if it cannot find the extension', async function () {
      const dotnetPath = 'road to nowhere';
      const executeCommandStub = stub(commands, 'executeCommand').returns(Promise.resolve({ dotnetPath }));
      const file = Uri.file(__filename);
      const testHtml = '<span>hello world</span>';
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns(testHtml);
      const disposeSpy = spy();
      const webview = { html: '' };
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview,
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
        get(key: string) {
          switch (key) {
            case 'fileFormatVersion':
              return '2010';
              break;
            case 'outPutFilePath':
              return undefined;
            default:
              break;
          }
        },
      } as unknown as WorkspaceConfiguration);
      const getExtensionStub = stub(extensions, 'getExtension').returns(undefined as unknown as Extension<unknown>);

      stubs.push(
        executeCommandStub,
        showErrorMessageStub,
        getWebviewContentStub,
        createWebviewPanelStub,
        getConfigurationStub,
        getExtensionStub,
      );

      await OOXMLValidator.validate(file);

      expect(disposeSpy.called).to.eq(true, 'panel.dispose() should have been called');
      expect(showErrorMessageStub.firstCall.firstArg).to.eq('Could not find OOXML Validate extension.');
    });

    test('should show an error and return if stderr has length', async function () {
      const testHtml = '<span>hello world</span>';
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns(testHtml);
      const disposeSpy = spy();
      const webview = { html: '' };
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview,
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
        get(key: string) {
          switch (key) {
            case 'fileFormatVersion':
              return '2010';
              break;
            case 'outPutFilePath':
              return undefined;
            default:
              break;
          }
        },
      } as unknown as WorkspaceConfiguration);
      const file = Uri.file(__filename);

      const extensionPath = 'foobar';
      const getExtensionStub = stub(extensions, 'getExtension').returns({ extensionPath } as Extension<unknown>);

      const dotnetPath = 'road to nowhere';
      const executeCommandStub = stub(commands, 'executeCommand').returns(Promise.resolve({ dotnetPath }));

      const errorMsg = 'oh nooooo what happened?';
      const spawnSyncStub = stub(child_process, 'spawnSync').returns({
        stdout: Buffer.from(JSON.stringify([])),
        stderr: Buffer.from(errorMsg),
        pid: 7,
        output: [null],
        status: 13,
        signal: null,
      });

      stubs.push(
        showErrorMessageStub,
        getWebviewContentStub,
        createWebviewPanelStub,
        getConfigurationStub,
        getExtensionStub,
        executeCommandStub,
        spawnSyncStub,
      );

      await OOXMLValidator.validate(file);

      expect(showErrorMessageStub.firstCall.firstArg).to.eq(`Failed to run OOXML Validator. The error was:\n${errorMsg}`);
    });

    test('should prompt the user to instal the dotnet runtime extension if trying to call it throws an error', async function () {
      const errMsg = 'dotnet.showAcquisitionLog command not found';
      const executeCommandStub = stub(commands, 'executeCommand').throws(errMsg);
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const file = Uri.file(__filename);

      stubs.push(executeCommandStub, showErrorMessageStub);

      await OOXMLValidator.validate(file);

      expect(showErrorMessageStub.firstCall.firstArg).to.eq(
        'The ".NET Install Tool for Extension Authors" VS Code extension\nMUST be installed for the OOXML Validator extension to work.',
      );
      expect(showErrorMessageStub.firstCall.lastArg).to.deep.eq({ modal: true });
    });

    test('should display an error if one is thrown', async function () {
      const errMsg = ['eek gads no tacos!!'];
      const executeCommandStub = stub(commands, 'executeCommand').throws(errMsg);
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const file = Uri.file(__filename);

      stubs.push(executeCommandStub, showErrorMessageStub);

      await OOXMLValidator.validate(file);

      expect(showErrorMessageStub.firstCall.firstArg).to.deep.eq(errMsg);
      expect(showErrorMessageStub.firstCall.lastArg).to.deep.eq({ modal: true });
    });
  });
});
