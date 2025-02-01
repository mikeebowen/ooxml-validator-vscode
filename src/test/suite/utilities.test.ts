import { expect } from 'chai';
import * as child_process from 'child_process';
import { join } from 'path';
import { SinonStub, stub } from 'sinon';
import { FileSystemError, window, workspace, WorkspaceConfiguration } from 'vscode';
import { ExtensionUtilities, logger, WindowUtilities, WorkspaceUtilities } from '../../utilities';

suite('Utilities', function () {
  this.timeout(15000);
  const stubs: SinonStub[] = [];

  teardown(function () {
    stubs.forEach(s => s.restore());
    stubs.length = 0;
  });

  suite('Extension Utilities', function () {
    suite('isDotNetRuntime', function () {
      test('should return true if the .NET runtime path is valid', async function () {
        const spawnSyncStub = stub(child_process, 'spawnSync').returns({
          stdout: Buffer.from('.NET runtime'),
          stderr: Buffer.from(''),
          pid: 7,
          output: [null],
          status: 13,
          signal: null,
        });
        const path = 'something, something, dark side';

        stubs.push(spawnSyncStub);

        const isRuntime = await ExtensionUtilities.isDotNetRuntime(path);

        expect(spawnSyncStub.calledOnce).to.be.true;
        expect(spawnSyncStub.calledWith(path, ['--info'])).to.be.true;
        expect(isRuntime).to.be.true;
      });

      test('should return false if the .NET runtim path is invalid', async function () {
        const spawnSyncStub = stub(child_process, 'spawnSync').returns({
          stdout: Buffer.from('Javascript runtime'),
          stderr: Buffer.from(''),
          pid: 7,
          output: [null],
          status: 13,
          signal: null,
        });
        const path = 'nacho nacho man';

        stubs.push(spawnSyncStub);

        const isRuntime = await ExtensionUtilities.isDotNetRuntime(path);

        expect(spawnSyncStub.calledOnce).to.be.true;
        expect(spawnSyncStub.calledWith(path, ['--info'])).to.be.true;
        expect(isRuntime).to.be.false;
      });
    });
  });

  suite('Window Utilities', function () {
    suite('showError', function () {
      test('should show an error message when the message is a string', async function () {
        const loggerStub = stub(logger, 'error');
        const showErrorMessageStub = stub(window, 'showErrorMessage');
        const err = 'mellow has been harshed';

        stubs.push(loggerStub, showErrorMessageStub);

        await WindowUtilities.showError(err);

        expect(loggerStub.calledOnce).to.be.true;
        expect(loggerStub.calledWith(err)).to.be.true;
        expect(showErrorMessageStub.calledOnce).to.be.true;
        expect(showErrorMessageStub.calledWith(err, { modal: false })).to.be.true;
      });

      test('should show an error message when the message is an error', async function () {
        const loggerStub = stub(logger, 'error');
        const showErrorMessageStub = stub(window, 'showErrorMessage');
        const err = new Error('ice cream is too cold');

        stubs.push(loggerStub, showErrorMessageStub);

        await WindowUtilities.showError(err, true);

        expect(loggerStub.calledOnce).to.be.true;
        expect(loggerStub.calledWith(err.message)).to.be.true;
        expect(showErrorMessageStub.calledOnce).to.be.true;
        expect(showErrorMessageStub.calledWith(err.message, { modal: true })).to.be.true;
      });
    });

    suite('showWarning', function () {
      test('should show a warning message', async function () {
        const loggerStub = stub(logger, 'warn');
        const showWarningMessageStub = stub(window, 'showWarningMessage');
        const message = 'the cake is a lie';
        const detail = 'but the pie is real';

        stubs.push(loggerStub, showWarningMessageStub);

        await WindowUtilities.showWarning(message, detail, true);

        expect(loggerStub.firstCall.firstArg).to.eq(message);
        expect(showWarningMessageStub.calledOnce).to.be.true;
        expect(showWarningMessageStub.calledWith(message, { modal: true, detail })).to.be.true;
      });
    });

    suite('createWebView', function () {
      test('should create a webview', function () {
        const createWebviewPanelStub = stub(window, 'createWebviewPanel');
        const viewType = 'view-type';
        const title = 'title';
        const showOptions = window.activeTextEditor?.viewColumn ?? 1;
        const options = { enableScripts: true };

        stubs.push(createWebviewPanelStub);

        WindowUtilities.createWebView(viewType, title, showOptions, options);

        expect(createWebviewPanelStub.calledOnce).to.be.true;
        expect(createWebviewPanelStub.calledWith(viewType, title, showOptions, options)).to.be.true;
      });
    });
  });

  suite('Workspace Utilities', function () {
    suite('createDirectory', function () {
      test('should create a directory', async function () {
        const loggerStub = stub(logger, 'trace');
        const directoryPath = join('c:', 'temp', 'directory');
        stubs.push(loggerStub);

        await WorkspaceUtilities.createDirectory(directoryPath);

        expect(loggerStub.firstCall.firstArg).to.eq(`Creating directory '${directoryPath}'`);
      });
    });

    suite('writeFile', function () {
      test('should write data to a file', async function () {
        const loggerStub = stub(logger, 'trace');
        const filePath = join('c:', 'temp', 'file.txt');
        const data = new Uint8Array(Buffer.from('hola mundo'));
        stubs.push(loggerStub);

        const res = await WorkspaceUtilities.writeFile(filePath, data);

        expect(res).to.be.true;
        expect(loggerStub.firstCall.firstArg).to.eq(`Writing file '${filePath}'`);
      });

      test('should return false for an unknown error', async function () {
        const loggerStub = stub(logger, 'error');
        const traceStub = stub(logger, 'trace').throws(new FileSystemError('unknown'));
        const filePath = join('c:', 'temp', 'file.txt');
        const data = new Uint8Array(Buffer.from('bonjour le monde'));

        stubs.push(traceStub, loggerStub);
        const res = await WorkspaceUtilities.writeFile(filePath, data);

        expect(res).to.be.false;
        expect(loggerStub.called).to.be.false;
      });

      test('should return false for an ebusy error', async function () {
        const loggerStub = stub(logger, 'error');
        const traceStub = stub(logger, 'trace').throws(new Error('ebusy error'));
        const filePath = join('c:', 'temp', 'file.txt');
        const data = new Uint8Array(Buffer.from('Hallo Welt'));

        stubs.push(traceStub, loggerStub);

        const res = await WorkspaceUtilities.writeFile(filePath, data);

        expect(res).to.be.false;
        expect(loggerStub.called).to.be.false;
      });

      test('should throw an error for other errors', async function () {
        const loggerStub = stub(logger, 'error');
        const traceStub = stub(logger, 'trace').throws(new Error('some other error'));
        const filePath = join('c:', 'temp', 'file.txt');
        const data = new Uint8Array(Buffer.from('Witaj świecie'));

        stubs.push(traceStub, loggerStub);

        try {
          await WorkspaceUtilities.writeFile(filePath, data);
        } catch (err: unknown) {
          const e = err as Error;
          expect(e.message).to.eq('some other error');
          expect(loggerStub.firstCall.firstArg).to.eq(`Error writing file '${filePath}'`);
        }
      });
    });

    suite('getConfigurationValue', function () {
      test('should get a configuration value', function () {
        const val = 'value';
        const getStub = stub().returns(val);
        const getConfigurationStub = stub(workspace, 'getConfiguration').returns({
          get: getStub,
        } as unknown as WorkspaceConfiguration);
        const section = 'section';
        const name = 'name';

        stubs.push(getConfigurationStub);

        const value = WorkspaceUtilities.getConfigurationValue(section, name);

        expect(getConfigurationStub.calledOnce).to.be.true;
        expect(getConfigurationStub.firstCall.firstArg).to.eq(section);
        expect(getStub.calledOnce).to.be.true;
        expect(getStub.firstCall.firstArg).to.eq(name);
        expect(value).to.eq('value');
      });
    });
  });
});