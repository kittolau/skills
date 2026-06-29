"use strict";
/*
 * pyfs.js
 *
 * Small helpers that mirror the subset of Python's pathlib / os behaviour used
 * by the ported scripts: glob, rglob, relative_to, resolve, iterdir, etc.
 *
 * Ordering: entries within a directory are sorted by name and the tree is
 * walked pre-order, which matches the alphabetical ordering pathlib yields on
 * NTFS (and keeps output deterministic across platforms).
 */

const fs = require("fs");
const path = require("path");

// Translate an fnmatch-style glob segment (*, ?, [seq]) to a RegExp.
function fnmatchToRegExp(pattern) {
  let re = "^";
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === "*") {
      re += "[^/]*";
    } else if (c === "?") {
      re += "[^/]";
    } else if (c === "[") {
      let j = i + 1;
      if (j < pattern.length && (pattern[j] === "!" || pattern[j] === "^")) j++;
      if (j < pattern.length && pattern[j] === "]") j++;
      while (j < pattern.length && pattern[j] !== "]") j++;
      if (j >= pattern.length) {
        re += "\\[";
      } else {
        let stuff = pattern.slice(i + 1, j);
        if (stuff.startsWith("!")) stuff = "^" + stuff.slice(1);
        re += "[" + stuff + "]";
        i = j;
      }
    } else {
      re += c.replace(/[.\\+^$(){}|=!<>:\-]/g, "\\$&");
    }
  }
  re += "$";
  return new RegExp(re);
}

function matchName(name, pattern) {
  return fnmatchToRegExp(pattern).test(name);
}

function listEntries(dir) {
  let entries;
  try {
    entries = fs.readdirSync(dir, { withFileTypes: true });
  } catch (e) {
    return [];
  }
  entries.sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
  return entries;
}

function exists(p) {
  try {
    fs.statSync(p);
    return true;
  } catch (e) {
    return false;
  }
}

function isFile(p) {
  try {
    return fs.statSync(p).isFile();
  } catch (e) {
    return false;
  }
}

function isDir(p) {
  try {
    return fs.statSync(p).isDirectory();
  } catch (e) {
    return false;
  }
}

// Path.iterdir(): immediate children (files and dirs), sorted by name.
function iterdir(dir) {
  return listEntries(dir).map((e) => path.join(dir, e.name));
}

/*
 * Path.glob(pattern): pattern may contain '/' separators. A single '*' wildcard
 * matches within one path segment. Walks segment by segment from base.
 * Returns matching paths (files or directories) in sorted, pre-order.
 */
function glob(base, pattern) {
  const segments = pattern.split("/").filter((s) => s.length > 0);
  let current = [base];
  for (let s = 0; s < segments.length; s++) {
    const seg = segments[s];
    const next = [];
    const isLast = s === segments.length - 1;
    for (const dir of current) {
      if (!isDir(dir)) continue;
      for (const entry of listEntries(dir)) {
        if (matchName(entry.name, seg)) {
          const full = path.join(dir, entry.name);
          if (isLast) {
            next.push(full);
          } else if (entry.isDirectory()) {
            next.push(full);
          }
        }
      }
    }
    current = next;
  }
  return current;
}

/*
 * Path.rglob(pattern): match `pattern` (no '/' expected here) against the
 * basename of every descendant (files and directories), pre-order, sorted.
 */
function rglob(base, pattern) {
  const results = [];
  function walk(dir) {
    for (const entry of listEntries(dir)) {
      const full = path.join(dir, entry.name);
      if (matchName(entry.name, pattern)) {
        results.push(full);
      }
      if (entry.isDirectory()) {
        walk(full);
      }
    }
  }
  if (isDir(base)) walk(base);
  return results;
}

// os.path.relpath / Path.relative_to -> OS-native separators (like str(Path)).
function relativeTo(base, p) {
  return path.relative(base, p);
}

// Python's str(Path) uses os.sep; expose a helper for clarity.
function toStr(p) {
  return p;
}

function resolve(p) {
  return path.resolve(p);
}

// Recursively copy a directory tree (mirrors shutil.copytree to a new dir).
function copytree(src, dest) {
  fs.mkdirSync(dest, { recursive: true });
  for (const entry of fs.readdirSync(src, { withFileTypes: true })) {
    const s = path.join(src, entry.name);
    const d = path.join(dest, entry.name);
    if (entry.isDirectory()) {
      copytree(s, d);
    } else if (entry.isSymbolicLink()) {
      fs.symlinkSync(fs.readlinkSync(s), d);
    } else {
      fs.copyFileSync(s, d);
    }
  }
}

module.exports = {
  fnmatchToRegExp,
  matchName,
  listEntries,
  exists,
  isFile,
  isDir,
  iterdir,
  glob,
  rglob,
  relativeTo,
  toStr,
  resolve,
  copytree,
};
