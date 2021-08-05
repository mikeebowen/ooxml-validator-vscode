'use strict';
const { rebuild } = require('electron-rebuild');

async function defaultTask(cb) {
  await rebuild({
    buildPath: __dirname,
    electronVersion: '13.1.7',
  });
  cb();
}

exports.default = defaultTask;
