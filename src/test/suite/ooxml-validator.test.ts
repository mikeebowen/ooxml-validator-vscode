import { window } from 'vscode';
import { SinonStub, stub } from 'sinon';
import { expect } from 'chai';
import * as path from 'path';

// You can import and use all API from the 'vscode' module
// as well as import your extension to test it
import * as vscode from 'vscode';
import OOXMLValidator from '../../ooxml-validator';

suite('OOXMLValidator', () => {
  const ooxmlValidator = new OOXMLValidator();
  const stubs: SinonStub[] = [];

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  test('createLogFile should throw an error if path is not absolute', async function () {
    const isAbsoluteStub = stub(path, 'isAbsolute').returns(false);
    const showErrorMessageStub = stub(window, 'showErrorMessage');
    stubs.push(isAbsoluteStub, showErrorMessageStub);
    OOXMLValidator.createLogFile([], 'tacocat');
    expect(isAbsoluteStub.calledWith('tacocat')).to.be.true;
    expect(showErrorMessageStub.calledOnce).to.be.true;
    return Promise.resolve();
  });
});
