const fs = require('fs');
const path = require('path');
const type = process.argv[2];

const vuePath = path.resolve(__dirname, '../node_modules/vue');
const vueBakPath = path.resolve(__dirname, '../node_modules/vue-bak');
if (type === 'bak') {
    if (!fs.existsSync(vueBakPath) && fs.existsSync(vuePath)) {
        fs.renameSync(vuePath, vueBakPath);
    }
} else {
    if (!fs.existsSync(vuePath) && fs.existsSync(vueBakPath)) {
        fs.renameSync(vueBakPath, vuePath);
    }
}
