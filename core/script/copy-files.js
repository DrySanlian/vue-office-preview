const fs = require('fs');
const path = require('path');

const pairs = process.argv.slice(2);
if (pairs.length === 0 || pairs.length % 2 !== 0) {
    throw new Error('Usage: node copy-files.js <source> <destination> [...]');
}

for (let index = 0; index < pairs.length; index += 2) {
    const source = path.resolve(process.cwd(), pairs[index]);
    const destination = path.resolve(process.cwd(), pairs[index + 1]);
    fs.mkdirSync(path.dirname(destination), { recursive: true });
    fs.copyFileSync(source, destination);
}
