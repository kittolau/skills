"use strict";
/*
 * Pack a directory into a DOCX, PPTX, or XLSX file.
 * (Port of pack.py)
 *
 * Validates with auto-repair, condenses XML formatting, and creates the Office file.
 *
 * Usage:
 *   node pack.js <input_directory> <output_file> [--original <file>] [--validate true|false]
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const AdmZip = require("adm-zip");
const minidom = require("../_pylib/minidom");
const pyfs = require("../_pylib/pyfs");
const {
  DOCXSchemaValidator,
  PPTXSchemaValidator,
  RedliningValidator,
} = require("./validators");

async function pack(
  inputDirectory,
  outputFile,
  originalFile = null,
  validate = true,
  inferAuthorFunc = null
) {
  const inputDir = inputDirectory;
  const outputPath = outputFile;
  const suffix = path.extname(outputPath).toLowerCase();

  if (!pyfs.isDir(inputDir)) {
    return [null, `Error: ${inputDir} is not a directory`];
  }

  if (![".docx", ".pptx", ".xlsx"].includes(suffix)) {
    return [null, `Error: ${outputFile} must be a .docx, .pptx, or .xlsx file`];
  }

  if (validate && originalFile) {
    const originalPath = originalFile;
    if (pyfs.exists(originalPath)) {
      const [success, output] = await _runValidation(inputDir, originalPath, suffix, inferAuthorFunc);
      if (output) console.log(output);
      if (!success) {
        return [null, `Error: Validation failed for ${inputDir}`];
      }
    }
  }

  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "pack-"));
  try {
    const tempContentDir = path.join(tempDir, "content");
    pyfs.copytree(inputDir, tempContentDir);

    for (const pattern of ["*.xml", "*.rels"]) {
      for (const xmlFile of pyfs.rglob(tempContentDir, pattern)) {
        _condenseXml(xmlFile);
      }
    }

    fs.mkdirSync(path.dirname(path.resolve(outputPath)), { recursive: true });

    const zip = new AdmZip();
    for (const f of pyfs.rglob(tempContentDir, "*")) {
      if (pyfs.isFile(f)) {
        const arc = path.relative(tempContentDir, f).split(path.sep).join("/");
        const dir = path.dirname(arc);
        zip.addFile(arc, fs.readFileSync(f));
        void dir;
      }
    }
    zip.writeZip(outputPath);
  } finally {
    try {
      fs.rmSync(tempDir, { recursive: true, force: true });
    } catch (e) {}
  }

  return [null, `Successfully packed ${inputDir} to ${outputFile}`];
}

async function _runValidation(unpackedDir, originalFile, suffix, inferAuthorFunc = null) {
  const outputLines = [];
  let validators = [];

  if (suffix === ".docx") {
    let author = "Claude";
    if (inferAuthorFunc) {
      try {
        author = inferAuthorFunc(unpackedDir, originalFile);
      } catch (e) {
        process.stderr.write(`Warning: ${e.message} Using default author 'Claude'.\n`);
      }
    }

    validators = [
      new DOCXSchemaValidator(unpackedDir, originalFile),
      new RedliningValidator(unpackedDir, originalFile, false, author),
    ];
  } else if (suffix === ".pptx") {
    validators = [new PPTXSchemaValidator(unpackedDir, originalFile)];
  }

  if (validators.length === 0) {
    return [true, null];
  }

  let totalRepairs = 0;
  for (const v of validators) totalRepairs += v.repair();
  if (totalRepairs) {
    outputLines.push(`Auto-repaired ${totalRepairs} issue(s)`);
  }

  let success = true;
  for (const v of validators) {
    // Mirror Python all(...) short-circuit semantics.
    if (!(await v.validate())) {
      success = false;
      break;
    }
  }

  if (success) {
    outputLines.push("All validations PASSED!");
  }

  return [success, outputLines.length ? outputLines.join("\n") : null];
}

function _condenseXml(xmlFile) {
  try {
    const dom = minidom.parseString(fs.readFileSync(xmlFile, "utf-8"));

    const all = dom.getElementsByTagName("*");
    for (let i = 0; i < all.length; i++) {
      const element = all[i];
      if (element.tagName.endsWith(":t")) continue;

      for (const child of minidom.childNodesArray(element).slice()) {
        if (
          (child.nodeType === minidom.TEXT_NODE &&
            child.nodeValue &&
            child.nodeValue.trim() === "") ||
          child.nodeType === minidom.COMMENT_NODE
        ) {
          element.removeChild(child);
        }
      }
    }

    fs.writeFileSync(xmlFile, minidom.toxml(dom, "UTF-8"));
  } catch (e) {
    process.stderr.write(`ERROR: Failed to parse ${path.basename(xmlFile)}: ${e.message}\n`);
    throw e;
  }
}

function parseArgs(argv) {
  const positional = [];
  const opts = { original: undefined, validate: true };
  for (let i = 0; i < argv.length; i++) {
    let arg = argv[i];
    if (arg === "--original" || arg.startsWith("--original=")) {
      if (arg.startsWith("--original=")) opts.original = arg.slice("--original=".length);
      else opts.original = argv[++i];
    } else if (arg === "--validate" || arg.startsWith("--validate=")) {
      let val;
      if (arg.startsWith("--validate=")) val = arg.slice("--validate=".length);
      else val = argv[++i];
      opts.validate = String(val).toLowerCase() === "true";
    } else {
      positional.push(arg);
    }
  }
  return { positional, opts };
}

async function main() {
  const { positional, opts } = parseArgs(process.argv.slice(2));
  const inputDirectory = positional[0];
  const outputFile = positional[1];

  const [, message] = await pack(inputDirectory, outputFile, opts.original || null, opts.validate);
  console.log(message);

  if (message.indexOf("Error") !== -1) {
    process.exit(1);
  }
}

module.exports = { pack, _runValidation, _condenseXml };

if (require.main === module) {
  main();
}
