"use strict";
/*
 * Command line tool to validate Office document XML files against XSD schemas
 * and tracked changes. (Port of validate.py)
 *
 * Usage:
 *   node validate.js <path> [--original <original_file>] [--auto-repair] [--author NAME]
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const AdmZip = require("adm-zip");
const pyfs = require("../_pylib/pyfs");
const {
  DOCXSchemaValidator,
  PPTXSchemaValidator,
  RedliningValidator,
} = require("./validators");

function assert(cond, msg) {
  if (!cond) {
    process.stderr.write(`AssertionError: ${msg}\n`);
    process.exit(1);
  }
}

function parseArgs(argv) {
  const opts = {
    path: undefined,
    original: null,
    verbose: false,
    auto_repair: false,
    author: "Claude",
  };
  const positional = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "--original" || arg.startsWith("--original=")) {
      opts.original = arg.startsWith("--original=") ? arg.slice("--original=".length) : argv[++i];
    } else if (arg === "-v" || arg === "--verbose") {
      opts.verbose = true;
    } else if (arg === "--auto-repair") {
      opts.auto_repair = true;
    } else if (arg === "--author" || arg.startsWith("--author=")) {
      opts.author = arg.startsWith("--author=") ? arg.slice("--author=".length) : argv[++i];
    } else {
      positional.push(arg);
    }
  }
  opts.path = positional[0];
  return opts;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));

  const p = args.path;
  assert(pyfs.exists(p), `Error: ${p} does not exist`);

  let originalFile = null;
  if (args.original) {
    originalFile = args.original;
    assert(pyfs.isFile(originalFile), `Error: ${originalFile} is not a file`);
    assert(
      [".docx", ".pptx", ".xlsx"].includes(path.extname(originalFile).toLowerCase()),
      `Error: ${originalFile} must be a .docx, .pptx, or .xlsx file`
    );
  }

  const fileExtension = path.extname(originalFile || p).toLowerCase();
  assert(
    [".docx", ".pptx", ".xlsx"].includes(fileExtension),
    `Error: Cannot determine file type from ${p}. Use --original or provide a .docx/.pptx/.xlsx file.`
  );

  let unpackedDir;
  if (pyfs.isFile(p) && [".docx", ".pptx", ".xlsx"].includes(path.extname(p).toLowerCase())) {
    const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "validate-"));
    const zip = new AdmZip(p);
    zip.extractAllTo(tempDir, true);
    unpackedDir = tempDir;
  } else {
    assert(pyfs.isDir(p), `Error: ${p} is not a directory or Office file`);
    unpackedDir = p;
  }

  let validators;
  switch (fileExtension) {
    case ".docx":
      validators = [new DOCXSchemaValidator(unpackedDir, originalFile, args.verbose)];
      if (originalFile) {
        validators.push(new RedliningValidator(unpackedDir, originalFile, args.verbose, args.author));
      }
      break;
    case ".pptx":
      validators = [new PPTXSchemaValidator(unpackedDir, originalFile, args.verbose)];
      break;
    default:
      console.log(`Error: Validation not supported for file type ${fileExtension}`);
      process.exit(1);
  }

  if (args.auto_repair) {
    let totalRepairs = 0;
    for (const v of validators) totalRepairs += v.repair();
    if (totalRepairs) {
      console.log(`Auto-repaired ${totalRepairs} issue(s)`);
    }
  }

  let success = true;
  for (const v of validators) {
    if (!(await v.validate())) {
      success = false;
      break;
    }
  }

  if (success) {
    console.log("All validations PASSED!");
  }

  process.exit(success ? 0 : 1);
}

if (require.main === module) {
  main();
}

module.exports = { main };
