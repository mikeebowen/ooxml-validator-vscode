/* eslint-disable @typescript-eslint/naming-convention */
import { Uri, ViewColumn, WebviewPanel, window, workspace, WorkspaceConfiguration } from 'vscode';
import { SinonStub, spy, stub, fake } from 'sinon';
import { expect } from 'chai';
import * as path from 'path';
import OOXMLValidator, { ValidationError, effEss, childProcess } from '../../ooxml-validator';
import { basename, normalize } from 'path';
import { TextEncoder } from 'util';
import * as csvWriter from 'csv-writer';
import { isEqual } from 'lodash';
import { PromiseWithChild } from 'child_process';
const edge = require('electron-edge-js');

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
      expect(isAbsoluteStub.calledWith('tacocat')).to.be.true;
      expect(showErrorMessageStub.calledOnce).to.be.true;
    });

    test('should create a json file if the file\'s extension is ".json"', function (done) {
      const testPath: string = normalize(path.join(__dirname, 'racecar.json'));
      const createDirectoryStub = stub(effEss, 'createDirectory');
      const writeFileStub: SinonStub = stub(effEss, 'writeFile');
      stubs.push(writeFileStub, createDirectoryStub);
      const errors: ValidationError[] = [new ValidationError({}), new ValidationError({})];
      OOXMLValidator.createLogFile(errors, testPath).then(() => {
        expect(createDirectoryStub.calledOnceWith(Uri.file(path.dirname(testPath)))).to.be.true;
        expect(writeFileStub.calledOnceWith(Uri.file(testPath), new TextEncoder().encode(JSON.stringify(errors)))).to.be.true;
        done();
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
            NamespacesDefinitions: {
              $values: ['firstNamespace', 'secondNamespace'],
            },
            Namespaces: {},
            PartUri: 'some/uri',
          },
        }),
        new ValidationError({
          Id: '2',
          Description: 'the second test error',
          Path: {
            NamespacesDefinitions: {
              $values: ['thirdNamespace', 'fourthNamespace'],
            },
            Namespaces: {},
            PartUri: 'some/other/uri',
          },
        }),
      ];
      const testHtml =
        '<!DOCTYPE html>\n        <html lang="en">\n        <head>\n            <meta charset="UTF-8">\n            <meta name="viewport" content="width=device-width, initial-scale=1.0">\n            <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>\n            <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>\n            <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>\n            <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>\n            <title>OOXML Validation Errors</title>\n            <body>\n              <div class="container-fluid pt-3 ol-3">\n              <div class="row pb-3">\n                <div class="col">\n                  <h1>There Were 2 Validation Errors Found</h1>\n  <h2>A log of these errors was saved as "/test-directory"</h2>\n                </div>\n              </div>\n              <div class="row pb-3">\n                <div class="col">\n                  <button\n                  class="btn btn-warn"\n                  type="button"\n                  data-toggle="collapse"\n                  data-target="#collapseExample"\n                  aria-expanded="false"\n                  aria-controls="collapseExample"\n                >\n                    View Errors\n                  </button>\n                </div>\n              </div>\n              <div class="row pb-3">\n                <div class="col">\n                  <div class="collapse" id="collapseExample">\n                    <div class="card card-body">\n                      <dl class="row">\n              <dt class="col-sm-3">Id</dt>\n              <dd class="col-sm-9">1</dd>\n              <dt class="col-sm-3">Description</dt>\n              <dd class="col-sm-9">the first test error</dd>\n              <dt class="col-sm-3">XPath</dt>\n              <dd class="col-sm-9">\n                undefined\n              </dd>\n              <dt class="col-sm-3">Part URI</dt>\n              <dd class="col-sm-9">some/uri</dd>\n              <dt class="col-sm-3">NamespacesDefinitions</dt>\n              <dd class="col-sm-9">\n                <ul>\n                  <li>firstNamespace</li><li>secondNamespace</li>\n                </ul>\n              </dd>\n            </dl><dl class="row">\n              <dt class="col-sm-3">Id</dt>\n              <dd class="col-sm-9">2</dd>\n              <dt class="col-sm-3">Description</dt>\n              <dd class="col-sm-9">the second test error</dd>\n              <dt class="col-sm-3">XPath</dt>\n              <dd class="col-sm-9">\n                undefined\n              </dd>\n              <dt class="col-sm-3">Part URI</dt>\n              <dd class="col-sm-9">some/other/uri</dd>\n              <dt class="col-sm-3">NamespacesDefinitions</dt>\n              <dd class="col-sm-9">\n                <ul>\n                  <li>thirdNamespace</li><li>fourthNamespace</li>\n                </ul>\n              </dd>\n            </dl>\n                    </div>\n                  </div>\n                </div>\n              </div>\n            </body>\n        </html>';
      const html = OOXMLValidator.getWebviewContent(validationErrors, 'test-file.csv', '/test-directory');
      expect(testHtml).to.equal(html);
      done();
    });

    test('should return the success html if there are no validation errors', function (done) {
      const validationErrors: ValidationError[] = [];
      const testHtml =
        '<!DOCTYPE html>\n      <html lang="en">\n      <head>\n          <meta charset="UTF-8">\n          <meta name="viewport" content="width=device-width, initial-scale=1.0">\n          <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>\n          <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>\n          <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>\n          <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>\n          <title>OOXML Validation Errors</title>\n          <body>\n            <div class="container pt-3 ol-3">\n              <div class="row">\n                <div class="col">\n                <div class="jumbotron">\n                <h1 class="display-4 text-center">No Errors Found!!</h1>\n                <p class="lead text-center">OOXML Validator did not find any validation errors in test-file.csv.</p>\n              </div>\n                </div>\n              </div>\n            </div>\n          </body>\n        </html>';
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
    test('should show error modal if exec returns an error', async function () {
      const execStub = stub(childProcess, 'execPromise').returns(
        Promise.resolve({ stdout: null, stderr: 'Out of coffee' }) as PromiseWithChild<any>,
      );
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      await OOXMLValidator.validate(Uri.file(__filename));
      expect(execStub.firstCall.firstArg).to.eq('dotnet --list-runtimes');
      expect(showErrorMessageStub.firstCall.firstArg).to.eq('Out of coffee');
      stubs.push(execStub, showErrorMessageStub);
    });

    test('should show an error modal if .Net Core 5 is not installed', async function () {
      const execStub = stub(childProcess, 'execPromise').returns(
        Promise.resolve({ stdout: 'Coffee! Coffee! Coffee!', stderr: null }) as PromiseWithChild<any>,
      );
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      await OOXMLValidator.validate(Uri.file(__filename));
      expect(showErrorMessageStub.firstCall.firstArg).to.eq(
        'OOXML Validator requires .Net 5 be installed.\nYou can download it from "https://dotnet.microsoft.com/download/dotnet"',
      );
      stubs.push(execStub, showErrorMessageStub);
    });

    test('should show an error modal if the edge callback is called with an error', async function () {
      const execStub = stub(childProcess, 'execPromise').returns(
        Promise.resolve({ stdout: 'NETCore.App 5.', stderr: null }) as PromiseWithChild<any>,
      );
      const showErrorMessageStub = stub(window, 'showErrorMessage');
      const getWebviewContentStub = stub(OOXMLValidator, 'getWebviewContent').returns('<span>hello world</span>');
      const disposeSpy = spy();
      const createWebviewPanelStub = stub(window, 'createWebviewPanel').returns({
        webview: { html: '' },
        dispose: disposeSpy,
      } as unknown as WebviewPanel);
      const getFake = fake.returns(undefined);
      const getConfigurationStub = stub(workspace, 'getConfiguration').returns({ get: getFake } as unknown as WorkspaceConfiguration);
      const file = Uri.file(__filename);
      const edgeStub = stub(edge, 'func').returns(function (data: any, callback: any) {
        expect(data).to.deep.eq(JSON.stringify({ fileName: file.fsPath, format: '4' }));
        callback('Need to walk the dog', null);
      });

      await OOXMLValidator.validate(file);
      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(true, 'panel.dispose() should have been called');
      expect(getFake.getCall(0).args[0]).to.eq('fileFormatVersion');
      expect(showErrorMessageStub.firstCall.firstArg).to.eq('Need to walk the dog');
      expect(showErrorMessageStub.firstCall.args[1]).to.deep.eq({ modal: true });
      stubs.push(execStub, showErrorMessageStub, getWebviewContentStub, createWebviewPanelStub, edgeStub, getConfigurationStub);
    });

    test('should show validation errors in the web view and use a different version of Office if one is assigned', async function () {
      const sdkValidationErrors = {
        $id: '1',
        $values: [
          {
            $id: '2023',
            Id: 'Sch_UnexpectedElementContentExpectingComplex',
            ErrorType: 0,
            Description:
              'The element has unexpected child element u0027http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMaxu0027.',
            Path: {
              $id: '2024',
              NamespacesDefinitions: {
                $id: '2025',
                $values: ['xmlns:c=u0022http://schemas.openxmlformats.org/drawingml/2006/chartu0022'],
              },
              Namespaces: {
                $id: '2026',
              },
              XPath: '/c:chartSpace[1]/c:chart[1]',
              PartUri: '/word/charts/chart1.xml',
            },
            Node: {
              $ref: '1252',
            },
            Part: {
              $ref: '1242',
            },
            RelatedNode: {
              $ref: '1588',
            },
            RelatedPart: null,
          },
          {
            $id: '2027',
            Id: 'Sch_UnexpectedElementContentExpectingComplex',
            ErrorType: 0,
            Description:
              'The element has unexpected child element u0027http://schemas.microsoft.com/office/drawing/2012/chart:leaderLinesu0027.',
            Path: {
              $id: '2028',
              NamespacesDefinitions: {
                $id: '2029',
                $values: ['xmlns:c=u0022http://schemas.openxmlformats.org/drawingml/2006/chartu0022'],
              },
              Namespaces: {
                $id: '2030',
              },
              XPath: '/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:scatterChart[1]/c:ser[2]/c:dLbls[1]/c:extLst[1]/c:ext[1]',
              PartUri: '/word/charts/chart1.xml',
            },
            Node: {
              $ref: '1425',
            },
            Part: {
              $ref: '1242',
            },
            RelatedNode: {
              $ref: '1427',
            },
            RelatedPart: null,
          },
          {
            $id: '2031',
            Id: 'Sch_UnexpectedElementContentExpectingComplex',
            ErrorType: 0,
            Description:
              'The element has unexpected child element u0027http://schemas.microsoft.com/office/drawing/2012/chart:leaderLinesu0027.',
            Path: {
              $id: '2032',
              NamespacesDefinitions: {
                $id: '2033',
                $values: ['xmlns:c=u0022http://schemas.openxmlformats.org/drawingml/2006/chartu0022'],
              },
              Namespaces: {
                $id: '2034',
              },
              XPath: '/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:scatterChart[1]/c:ser[1]/c:dLbls[1]/c:extLst[1]/c:ext[1]',
              PartUri: '/word/charts/chart1.xml',
            },
            Node: {
              $ref: '1320',
            },
            Part: {
              $ref: '1242',
            },
            RelatedNode: {
              $ref: '1322',
            },
            RelatedPart: null,
          },
        ],
      };
      const validationErrors = sdkValidationErrors.$values.map((v: any) => new ValidationError(v));
      const testHtml = '<span>hello world</span>';
      const execStub = stub(childProcess, 'execPromise').returns(
        Promise.resolve({ stdout: 'NETCore.App 5.', stderr: null }) as PromiseWithChild<any>,
      );
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
      const edgeStub = stub(edge, 'func').returns(function (data: any, callback: any) {
        expect(data).to.deep.eq(JSON.stringify({ fileName: file.fsPath, format: '1' }));
        callback(null, JSON.stringify(sdkValidationErrors));
      });
      //OOXMLValidator.getWebviewContent(validationErrors, basename(file.fsPath), path)
      await OOXMLValidator.validate(file);
      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(false, 'panel.dispose() should not have been called');
      // expect(getFake.getCall(0).args[1]).to.eq('outPutFilePath');
      expect(showErrorMessageStub.callCount).to.eq(0);
      expect(getWebviewContentStub.getCall(1).firstArg).to.deep.eq(validationErrors);
      expect(getWebviewContentStub.getCall(1).args[1]).to.eq(basename(file.fsPath));
      expect(getWebviewContentStub.getCall(1).args[2]).to.be.undefined;
      expect(webview.html).to.eq(testHtml);
      stubs.push(execStub, showErrorMessageStub, getWebviewContentStub, createWebviewPanelStub, edgeStub, getConfigurationStub);
    });

    // eslint-disable-next-line max-len
    test('should show validation errors in the web view, create a file with the contents and use a different version of Office if one is assigned', async function () {
      const sdkValidationErrors = {
        $id: '1',
        $values: [
          {
            $id: '2023',
            Id: 'Sch_UnexpectedElementContentExpectingComplex',
            ErrorType: 0,
            Description:
              'The element has unexpected child element u0027http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMaxu0027.',
            Path: {
              $id: '2024',
              NamespacesDefinitions: {
                $id: '2025',
                $values: ['xmlns:c=u0022http://schemas.openxmlformats.org/drawingml/2006/chartu0022'],
              },
              Namespaces: {
                $id: '2026',
              },
              XPath: '/c:chartSpace[1]/c:chart[1]',
              PartUri: '/word/charts/chart1.xml',
            },
            Node: {
              $ref: '1252',
            },
            Part: {
              $ref: '1242',
            },
            RelatedNode: {
              $ref: '1588',
            },
            RelatedPart: null,
          },
          {
            $id: '2027',
            Id: 'Sch_UnexpectedElementContentExpectingComplex',
            ErrorType: 0,
            Description:
              'The element has unexpected child element u0027http://schemas.microsoft.com/office/drawing/2012/chart:leaderLinesu0027.',
            Path: {
              $id: '2028',
              NamespacesDefinitions: {
                $id: '2029',
                $values: ['xmlns:c=u0022http://schemas.openxmlformats.org/drawingml/2006/chartu0022'],
              },
              Namespaces: {
                $id: '2030',
              },
              XPath: '/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:scatterChart[1]/c:ser[2]/c:dLbls[1]/c:extLst[1]/c:ext[1]',
              PartUri: '/word/charts/chart1.xml',
            },
            Node: {
              $ref: '1425',
            },
            Part: {
              $ref: '1242',
            },
            RelatedNode: {
              $ref: '1427',
            },
            RelatedPart: null,
          },
          {
            $id: '2031',
            Id: 'Sch_UnexpectedElementContentExpectingComplex',
            ErrorType: 0,
            Description:
              'The element has unexpected child element u0027http://schemas.microsoft.com/office/drawing/2012/chart:leaderLinesu0027.',
            Path: {
              $id: '2032',
              NamespacesDefinitions: {
                $id: '2033',
                $values: ['xmlns:c=u0022http://schemas.openxmlformats.org/drawingml/2006/chartu0022'],
              },
              Namespaces: {
                $id: '2034',
              },
              XPath: '/c:chartSpace[1]/c:chart[1]/c:plotArea[1]/c:scatterChart[1]/c:ser[1]/c:dLbls[1]/c:extLst[1]/c:ext[1]',
              PartUri: '/word/charts/chart1.xml',
            },
            Node: {
              $ref: '1320',
            },
            Part: {
              $ref: '1242',
            },
            RelatedNode: {
              $ref: '1322',
            },
            RelatedPart: null,
          },
        ],
      };
      const validationErrors = sdkValidationErrors.$values.map((v: any) => new ValidationError(v));
      const testHtml = '<span>hello world</span>';
      const testFilePath = 'C:\\source\\test\\errors\\tacocat.csv';
      const execStub = stub(childProcess, 'execPromise').returns(
        Promise.resolve({ stdout: 'NETCore.App 5.', stderr: null }) as PromiseWithChild<any>,
      );
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
      const edgeStub = stub(edge, 'func').returns(function (data: any, callback: any) {
        expect(data).to.deep.eq(JSON.stringify({ fileName: file.fsPath, format: '2' }));
        callback(null, JSON.stringify(sdkValidationErrors));
      });
      const createLogFileStub = stub(OOXMLValidator, 'createLogFile');
      await OOXMLValidator.validate(file);
      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(false, 'panel.dispose() should not have been called');
      expect(getWebviewContentStub.getCall(1).firstArg).to.deep.eq(validationErrors);
      expect(getWebviewContentStub.getCall(1).args[1]).to.eq(basename(file.fsPath));
      expect(getWebviewContentStub.getCall(1).args[2]).to.eq(testFilePath);
      expect(createLogFileStub.firstCall.firstArg).to.deep.eq(validationErrors);
      expect(createLogFileStub.firstCall.args[1]).to.eq(testFilePath);
      expect(webview.html).to.eq(testHtml);
      stubs.push(
        execStub,
        showErrorMessageStub,
        getWebviewContentStub,
        createWebviewPanelStub,
        edgeStub,
        getConfigurationStub,
        createLogFileStub,
      );
    });

    test('should show the no errors view if there are no validation errors', async function () {
      const sdkValidationErrors = {
        $id: '1',
        $values: [],
      };
      const validationErrors = sdkValidationErrors.$values.map((v: any) => new ValidationError(v));
      const testHtml = '<span>hello world</span>';
      const execStub = stub(childProcess, 'execPromise').returns(
        Promise.resolve({ stdout: 'NETCore.App 5.', stderr: null }) as PromiseWithChild<any>,
      );
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
      const edgeStub = stub(edge, 'func').returns(function (data: any, callback: any) {
        expect(data).to.deep.eq(JSON.stringify({ fileName: file.fsPath, format: '1' }));
        callback(null, JSON.stringify(sdkValidationErrors));
      });

      await OOXMLValidator.validate(file);
      expect(createWebviewPanelStub.firstCall.firstArg).to.eq('validateOOXML');
      expect(createWebviewPanelStub.firstCall.args[1]).to.eq('OOXML Validate');
      expect(createWebviewPanelStub.firstCall.args[2]).to.deep.eq(ViewColumn.One);
      expect(createWebviewPanelStub.firstCall.args[3]).to.deep.eq({ enableScripts: true });
      expect(disposeSpy.called).to.eq(false, 'panel.dispose() should not have been called');
      expect(showErrorMessageStub.callCount).to.eq(0);
      expect(getWebviewContentStub.getCall(1).firstArg).to.deep.eq(validationErrors);
      expect(getWebviewContentStub.getCall(1).args[1]).to.eq(basename(file.fsPath));
      expect(getWebviewContentStub.getCall(1).args[2]).to.be.undefined;
      expect(webview.html).to.eq(testHtml);
      stubs.push(execStub, showErrorMessageStub, getWebviewContentStub, createWebviewPanelStub, edgeStub, getConfigurationStub);
    });
  });
});
