// Mac: give the app an ad-hoc signature so it opens on Apple Silicon without a paid certificate.
// (Without this, macOS calls an unsigned download "damaged".) Windows: nothing to do.
const { execFileSync } = require('child_process');
const path = require('path');
exports.default = async function (ctx) {
  if (ctx.electronPlatformName !== 'darwin') { return; }
  const app = path.join(ctx.appOutDir, ctx.packager.appInfo.productFilename + '.app');
  execFileSync('codesign', ['--force', '--deep', '--sign', '-', app], { stdio: 'inherit' });
};
