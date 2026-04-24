// Tiny JSON-backed config for things Ketcher Desktop remembers between
// runs (currently just the Excel catalog path). Lives at
// $userData/config.json.
//
// Design note: we deliberately avoid electron-store or conf to keep the
// dependency graph small. This file's needs are trivial — one object,
// synchronous-looking API, no schema validation.

'use strict';

const fs = require('fs/promises');
const path = require('path');
const { app } = require('electron');

function configPath() {
  return path.join(app.getPath('userData'), 'config.json');
}

async function load() {
  try {
    const raw = await fs.readFile(configPath(), 'utf8');
    const obj = JSON.parse(raw);
    return obj && typeof obj === 'object' ? obj : {};
  } catch {
    return {};
  }
}

async function save(cfg) {
  const file = configPath();
  await fs.mkdir(path.dirname(file), { recursive: true });
  await fs.writeFile(file, JSON.stringify(cfg, null, 2), 'utf8');
}

// Get a single value.
async function get(key) {
  const cfg = await load();
  return cfg[key];
}

// Set a single value (shallow merge).
async function set(key, value) {
  const cfg = await load();
  cfg[key] = value;
  await save(cfg);
}

module.exports = { load, save, get, set, path: configPath };
