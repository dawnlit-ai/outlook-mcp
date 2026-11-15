// What the test files share. Not a test itself: the runner takes `*.test.js`.
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

/** Every directory `tempPath` made, removed when the test process exits. */
const scratchDirs = [];
process.on('exit', () => {
    for (const dir of scratchDirs) fs.rmSync(dir, {recursive: true, force: true});
});

/** A path for `name` in a fresh directory, which is removed when the test process exits. */
function tempPath(name) {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-test-'));
    scratchDirs.push(dir);
    return path.join(dir, name);
}

/** A file holding `contents`, in a directory of its own. */
function writeTemp(name, contents) {
    const file = tempPath(name);
    fs.writeFileSync(file, contents);
    return file;
}

module.exports = {tempPath, writeTemp};
