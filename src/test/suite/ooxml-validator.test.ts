import { Uri, window, workspace } from 'vscode';
import { SinonStub, spy, stub } from 'sinon';
import { expect } from 'chai';
import * as path from 'path';
import OOXMLValidator, { ValidationError, effEss } from '../../ooxml-validator';
import { normalize } from 'path';
import { TextEncoder } from 'util';
import * as csvWriter from 'csv-writer';
import { isEqual } from 'lodash';

suite('OOXMLValidator', function () {
  this.timeout(15000);
  const ooxmlValidator = new OOXMLValidator();
  const stubs: SinonStub[] = [];

  test('createLogFile should throw an error if path is not absolute', async function () {
    const isAbsoluteStub = stub(path, 'isAbsolute').returns(false);
    const showErrorMessageStub = stub(window, 'showErrorMessage').returns(Promise.resolve() as Thenable<undefined>);
    stubs.push(isAbsoluteStub, showErrorMessageStub);
    await OOXMLValidator.createLogFile([], 'tacocat');
    expect(isAbsoluteStub.calledWith('tacocat')).to.be.true;
    expect(showErrorMessageStub.calledOnce).to.be.true;
  });

  test('createLogFile should create a json file if the file\'s extension is ".json"', function (done) {
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

  test('createLogFile should create a csv file if json is not specified', async function () {
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

  test('getWebViewContent should return the errors in html if there are validation errors', function (done) {
    const validationErrors = [
      // eslint-disable-next-line @typescript-eslint/naming-convention
      new ValidationError({ Id: '1', Description: 'the first test error' }),
      // eslint-disable-next-line @typescript-eslint/naming-convention
      new ValidationError({ Id: '2', Description: 'the second test error' }),
    ];
    const testHtml =
      '<!DOCTYPE html>\n        <html lang="en">\n        <head>\n            <meta charset="UTF-8">\n            <meta name="viewport" content="width=device-width, initial-scale=1.0">\n            <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>\n            <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>\n            <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>\n            <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>\n            <title>OOXML Validation Errors</title>\n            <body>\n              <div class="container-fluid pt-3 ol-3">\n              <div class="row pb-3">\n                <div class="col">\n                  <h1>There Were 2 Validation Errors Found</h1>\n  <h2>A log of these errors was saved as "/test-directory"</h2>\n                </div>\n              </div>\n              <div class="row pb-3">\n                <div class="col">\n                  <button class="btn btn-warn" type="button" data-toggle="collapse" data-target="#collapseExample" aria-expanded="false" aria-controls="collapseExample">\n                    View Errors\n                  </button>\n                </div>\n              </div>\n              <div class="row pb-3">\n                <div class="col">\n                  <div class="collapse" id="collapseExample">\n                    <div class="card card-body">\n                      <dl class="row">\n              <dt class="col-sm-3">Id</dt>\n              <dd class="col-sm-9">1</dd>\n              <dt class="col-sm-3">Description</dt>\n              <dd class="col-sm-9">the first test error</dd>\n              <dt class="col-sm-3">XPath</dt>\n              <dd class="col-sm-9">\n                undefined\n              </dd>\n              <dt class="col-sm-3">Part URI</dt>\n              <dd class="col-sm-9">undefined</dd>\n              <dt class="col-sm-3">NamespacesDefinitions</dt>\n              <dd class="col-sm-9">\n                <ul>\n                  undefined\n                </ul>\n              </dd>\n            </dl><dl class="row">\n              <dt class="col-sm-3">Id</dt>\n              <dd class="col-sm-9">2</dd>\n              <dt class="col-sm-3">Description</dt>\n              <dd class="col-sm-9">the second test error</dd>\n              <dt class="col-sm-3">XPath</dt>\n              <dd class="col-sm-9">\n                undefined\n              </dd>\n              <dt class="col-sm-3">Part URI</dt>\n              <dd class="col-sm-9">undefined</dd>\n              <dt class="col-sm-3">NamespacesDefinitions</dt>\n              <dd class="col-sm-9">\n                <ul>\n                  undefined\n                </ul>\n              </dd>\n            </dl>\n                    </div>\n                  </div>\n                </div>\n              </div>\n            </body>\n        </html>';
    const html = OOXMLValidator.getWebviewContent(validationErrors, 'test-file.csv', '/test-directory');
    expect(testHtml).to.equal(html);
    done();
  });

  test('getWebViewContent should return the success html if there are no validation errors', function (done) {
    const validationErrors: ValidationError[] = [];
    const testHtml =
      '<!DOCTYPE html>\n      <html lang="en">\n      <head>\n          <meta charset="UTF-8">\n          <meta name="viewport" content="width=device-width, initial-scale=1.0">\n          <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/css/bootstrap.min.css" integrity="sha384-Gn5384xqQ1aoWXA+058RXPxPg6fy4IWvTNh0E263XmFcJlSAwiGgFAW/dAiS6JXm" crossorigin="anonymous"></head>\n          <script src="https://code.jquery.com/jquery-3.2.1.slim.min.js" integrity="sha384-KJ3o2DKtIkvYIK3UENzmM7KCkRr/rE9/Qpg6aAZGJwFDMVNA/GpGFF93hXpG5KkN" crossorigin="anonymous"></script>\n          <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.12.9/umd/popper.min.js" integrity="sha384-ApNbgh9B+Y1QKtv3Rn7W3mgPxhU9K/ScQsAP7hUibX39j7fakFPskvXusvfa0b4Q" crossorigin="anonymous"></script>\n          <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0/js/bootstrap.min.js" integrity="sha384-JZR6Spejh4U02d8jOt6vLEHfe/JQGiRRSQQxSfFWpi1MquVdAyjUar5+76PVCmYl" crossorigin="anonymous"></script>\n          <title>OOXML Validation Errors</title>\n          <body>\n            <div class="container pt-3 ol-3">\n              <div class="row">\n                <div class="col">\n                <div class="jumbotron">\n                <h1 class="display-4 text-center">No Open Office XML Validation Errors Found!!</h1>\n                <p class="lead text-center">OOXML Validator did not find any validation errors in test-file.csv.</p>\n              </div>\n                </div>\n              </div>\n            </div>\n          </body>\n        </html>';
    const html = OOXMLValidator.getWebviewContent(validationErrors, 'test-file.csv', '/test-directory');
    expect(testHtml).to.equal(html);
    done();
  });

  test('getWebViewContent should return the loading html if the validation errors parameter is undefined ', function (done) {
    const testHtml =
      '<!DOCTYPE html>\n            <html lang="en">\n            <head>\n                <meta charset="UTF-8">\n                <meta name="viewport" content="width=device-width, initial-scale=1.0">\n                <style>\n                .container {\n                  display: -webkit-box;\n                  display: -ms-flexbox;\n                  display: flex;\n                  -webkit-box-align: center;\n                      -ms-flex-align: center;\n                          align-items: center;\n                  -webkit-box-pack: center;\n                      -ms-flex-pack: center;\n                          justify-content: center;\n                  min-height: 100vh;\n                  background-color: #ededed;\n                }\n\n                .loader {\n                  max-width: 15rem;\n                  width: 100%;\n                  height: auto;\n                  stroke-linecap: round;\n                }\n\n                circle {\n                  fill: none;\n                  stroke-width: 3.5;\n                  -webkit-animation-name: preloader;\n                          animation-name: preloader;\n                  -webkit-animation-duration: 3s;\n                          animation-duration: 3s;\n                  -webkit-animation-iteration-count: infinite;\n                          animation-iteration-count: infinite;\n                  -webkit-animation-timing-function: ease-in-out;\n                          animation-timing-function: ease-in-out;\n                  -webkit-transform-origin: 170px 170px;\n                          transform-origin: 170px 170px;\n                  will-change: transform;\n                }\n                circle:nth-of-type(1) {\n                  stroke-dasharray: 550;\n                }\n                circle:nth-of-type(2) {\n                  stroke-dasharray: 500;\n                }\n                circle:nth-of-type(3) {\n                  stroke-dasharray: 450;\n                }\n                circle:nth-of-type(4) {\n                  stroke-dasharray: 300;\n                }\n                circle:nth-of-type(1) {\n                  -webkit-animation-delay: -0.15s;\n                          animation-delay: -0.15s;\n                }\n                circle:nth-of-type(2) {\n                  -webkit-animation-delay: -0.3s;\n                          animation-delay: -0.3s;\n                }\n                circle:nth-of-type(3) {\n                  -webkit-animation-delay: -0.45s;\n                  -moz-animation-delay:  -0.45s;\n                          animation-delay: -0.45s;\n                }\n                circle:nth-of-type(4) {\n                  -webkit-animation-delay: -0.6s;\n                  -moz-animation-delay: -0.6s;\n                          animation-delay: -0.6s;\n                }\n\n                @-webkit-keyframes preloader {\n                  50% {\n                    -webkit-transform: rotate(360deg);\n                            transform: rotate(360deg);\n                  }\n                }\n\n                @keyframes preloader {\n                  50% {\n                    -webkit-transform: rotate(360deg);\n                            transform: rotate(360deg);\n                  }\n                }\n                </style>\n                <title>OOXML Validation Errors</title>\n            </head>\n            <body>\n            <div class="container">\n            <div style="display: block">\n              <div style="display: block">\n                 <h1>Validating OOXML File</h1>\n              </div>\n                  <svg class="loader" xmlns="http://www.w3.org/2000/svg" viewBox="0 0 340 340">\n                    <circle cx="170" cy="170" r="160" stroke="#E2007C"/>\n                    <circle cx="170" cy="170" r="135" stroke="#404041"/>\n                    <circle cx="170" cy="170" r="110" stroke="#E2007C"/>\n                    <circle cx="170" cy="170" r="85" stroke="#404041"/>\n                  </svg>\n                </div>\n\n          </div>\n            </body>\n            </html>';
    const html = OOXMLValidator.getWebviewContent();
    expect(testHtml).to.equal(html);
    done();
  });

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });
});
