const fs = require('fs');
const path = require('path');

function stripTestOnlyBlocks(file) {
  let content = fs.readFileSync(file, 'utf8');
  // Remove blocks between #TEST-ONLY-START and #TEST-ONLY-END, inclusive
  content = content.replace(/#TEST-ONLY-START[\s\S]*?#TEST-ONLY-END/g, '');
  fs.writeFileSync(file, content, 'utf8');
}

function walk(dir) {
  fs.readdirSync(dir).forEach(file => {
    const fullPath = path.join(dir, file);
    if (fs.statSync(fullPath).isDirectory()) {
      walk(fullPath);
    } else if (fullPath.endsWith('.ts')) {
      stripTestOnlyBlocks(fullPath);
    }
  });
}

walk(path.resolve(__dirname, '../dist'));