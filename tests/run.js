// Runner que carrega todos os *.test.js e roda.
const fs = require('fs');
const path = require('path');
const { run } = require('./harness');

const dir = __dirname;
fs.readdirSync(dir).filter(f => f.endsWith('.test.js')).sort().forEach(f => {
  console.log('\n## ' + f);
  require(path.join(dir, f));
});

run();
