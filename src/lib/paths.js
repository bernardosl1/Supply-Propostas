const path = require('node:path');

function getPropostasDir() {
  return process.env.SUPPLY_MARINE_PROPOSTAS_DIR
    || path.resolve(__dirname, '..', '..', 'propostas');
}

function getIndexPath() {
  return path.join(getPropostasDir(), 'index.json');
}

module.exports = {
  getPropostasDir,
  getIndexPath
};
