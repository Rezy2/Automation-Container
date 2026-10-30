const fs = require('fs');
const path = require('path');

const root = path.resolve(__dirname, '..');
const files = ['src/config.gs', 'src/parser.gs', 'src/material.gs', 'src/draft.gs', 'src/fonnte.gs', 'src/setup.gs', 'src/dummy.gs'];
const bundle = files
  .map(function (file) {
    return fs.readFileSync(path.join(root, file), 'utf8');
  })
  .join('\n\n');

fs.writeFileSync(path.join(root, 'setup.gs'), bundle);
console.log('Built setup.gs from ' + files.length + ' source files.');
