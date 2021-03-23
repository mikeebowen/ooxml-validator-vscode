import { Uri, window, workspace } from 'vscode';
import { SinonStub, stub } from 'sinon';
import { expect } from 'chai';
import * as path from 'path';
import OOXMLValidator, { ValidationError } from '../../ooxml-validator';
import { normalize } from 'path';
import { TextEncoder } from 'util';
import isEqual = require('lodash.isequal');

suite('OOXMLValidator', () => {
  const ooxmlValidator = new OOXMLValidator();
  const stubs: SinonStub[] = [];

  test('createLogFile should throw an error if path is not absolute', async function () {
    const isAbsoluteStub = stub(path, 'isAbsolute').returns(false);
    const showErrorMessageStub = stub(window, 'showErrorMessage').returns(Promise.resolve() as Thenable<undefined>);
    stubs.push(isAbsoluteStub, showErrorMessageStub);
    await OOXMLValidator.createLogFile([], 'tacocat');
    expect(isAbsoluteStub.calledWith('tacocat')).to.be.true;
    expect(showErrorMessageStub.calledOnce).to.be.true;
    return Promise.resolve();
  });
  test.only('createLogFile should create a json file if the file\'s extension is ".json"', async function () {
    const testPath: string = normalize(path.join(__dirname, 'racecar.json'));
    const createDirectoryStub = stub(workspace.fs, 'createDirectory').callsFake(async (uri: Uri) => {
      expect(isEqual(uri, Uri.file(path.dirname(testPath)))).to.be.true;
      return Promise.resolve();
    });
    const writeFileStub: SinonStub = stub(workspace.fs, 'writeFile');
    stubs.push(createDirectoryStub, writeFileStub);
    const errors: ValidationError[] = [new ValidationError({}), new ValidationError({})];
    await OOXMLValidator.createLogFile(errors, testPath);
    expect(writeFileStub.calledOnceWith(Uri.file(testPath), new TextEncoder().encode(JSON.stringify(errors)))).to.be.true;
    return Promise.resolve();
  });
  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });
});
