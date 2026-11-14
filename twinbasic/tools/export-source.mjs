import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { exportProject } from './impexp.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const project = path.join(root, 'StardiskPlayer2NoWMP.twinproj');
const source = path.join(root, 'source');

// Never follow links while synchronizing the versioned source directory.
function list(folder, prefix = '') {
  if (fs.lstatSync(folder).isSymbolicLink()) throw new Error(`Symbolic link: ${folder}`);
  return fs.readdirSync(folder, { withFileTypes: true }).flatMap(entry => {
    const rel = prefix ? `${prefix}/${entry.name}` : entry.name;
    const full = path.join(folder, entry.name);
    if (entry.isSymbolicLink()) throw new Error(`Symbolic link: ${full}`);
    return entry.isDirectory() ? list(full, rel) : [rel];
  });
}

function removeEmptyFolders(folder) {
  for (const entry of fs.readdirSync(folder, { withFileTypes: true })) {
    if (!entry.isDirectory()) continue;
    const full = path.join(folder, entry.name);
    removeEmptyFolders(full);
    if (fs.readdirSync(full).length === 0) fs.rmdirSync(full);
  }
}

const temporaryRoot = fs.realpathSync(os.tmpdir());
const temporary = fs.mkdtempSync(path.join(temporaryRoot, 'stardisk-twinbasic-'));
try {
  const result = exportProject(project, temporary);
  if (result.repeated?.length) throw new Error('Duplicate entries in project; source was not updated.');
  const incoming = list(temporary);
  if (!incoming.includes('Settings')) throw new Error('Export has no Settings file.');
  const previous = fs.existsSync(source) ? list(source) : [];
  const keep = new Set(incoming);
  // Export into an empty directory first: CLI export alone retains deleted files.
  for (const rel of previous) {
    if (!keep.has(rel)) fs.unlinkSync(path.join(source, rel));
  }
  if (fs.existsSync(source)) removeEmptyFolders(source);
  for (const rel of incoming) {
    const target = path.join(source, rel);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.copyFileSync(path.join(temporary, rel), target);
  }
  console.log(`Exported ${incoming.length} files to ${source}. Review git diff before committing.`);
} finally {
  // Only remove the private temporary directory created by this invocation.
  if (path.dirname(temporary) !== temporaryRoot || !path.basename(temporary).startsWith('stardisk-twinbasic-')) {
    throw new Error('Refusing to remove an unexpected temporary path.');
  }
  fs.rmSync(temporary, { recursive: true, force: true });
}
