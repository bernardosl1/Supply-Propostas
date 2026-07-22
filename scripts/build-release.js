const { Arch, build, Platform } = require('electron-builder');

process.noAsar = true;

build({
  targets: Platform.WINDOWS.createTarget(['nsis'], Arch.x64),
  publish: 'never'
}).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
