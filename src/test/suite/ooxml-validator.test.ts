import { Uri, window, workspace } from 'vscode';
import { createStubInstance, SinonStub, spy, stub, SinonSpy } from 'sinon';
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

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });
});
