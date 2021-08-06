'use strict';
const { rebuild } = require('electron-rebuild');
const { exec } = require('child_process');
const got = require('got');
const { promisify } = require('util');
const execPromise = promisify(exec);

async function defaultTask(cb) {
  try {
    const { stdout: version } = await execPromise('code --version');
    const data = await got(`https://raw.githubusercontent.com/microsoft/vscode/${version.split('\n')[0]}/.yarnrc`);
    const electronVersion = data.body.split('\n')[1].split(' ')[1].replace(/"/g, '');

    await rebuild({
      buildPath: __dirname,
      electronVersion,
    });
    cb();
  } catch (error) {
    console.error(error.response || error);
  }
}

exports.default = defaultTask;
