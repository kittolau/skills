"use strict";
/*
 * Unpack Office files (DOCX, PPTX, XLSX) for editing. (Port of unpack.py)
 *
 * Usage: node unpack.js <office_file> <output_dir> [options]
 */

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const minidom = require("../_pylib/minidom");
const pyfs = require("../_pylib/pyfs");
const { mergeRuns: doMergeRuns } = require("./helpers/merge_runs");
const { simplifyRedlines: doSimplifyRedlines } = require("./helpers/simplify_redlines");

const SMART_QUOTE_REPLACEMENTS = {
  "\u201c": "&#x201C;",
  "\u201d": "&#x201D;",
  "\u2018": "&#x2018;",
  "\u2019": "&#x2019;",
};

function unpack(inputFile, outputDirectory, mergeRuns, simplifyRedlines) {
  if (mergeRuns === undefined) mergeRuns = true;
  if (simplifyRedlines === undefined) simplifyRedlines = true;

  const suffix = path.extname(inputFile).toLowerCase();

  if (!pyfs.exists(inputFile)) {
    return [null, `Error: ${inputFile} does not exist`];
  }

  if (![".docx", ".pptx", ".xlsx"].includes(suffix)) {
    return [null, `Error: ${inputFile} must be a .docx, .pptx, or .xlsx file`];
  }

  let zipFailed = false;
  try {
    fs.mkdirSync(outputDirectory, { recursive: true });

    let zip;
    try {
      zip = new AdmZip(inputFile);
      zip.extractAllTo(outputDirectory, true);
    } catch (e) {
      zipFailed = true;
      throw e;
    }

    const xmlFiles = pyfs
      .rglob(outputDirectory, "*.xml")
      .concat(pyfs.rglob(outputDirectory, "*.rels"));
    for (const xmlFile of xmlFiles) {
      prettyPrintXml(xmlFile);
    }

    let message = `Unpacked ${inputFile} (${xmlFiles.length} XML files)`;

    if (suffix === ".docx") {
      if (simplifyRedlines) {
        const [simplifyCount] = doSimplifyRedlines(outputDirectory);
        message += `, simplified ${simplifyCount} tracked changes`;
      }

      if (mergeRuns) {
        const [mergeCount] = doMergeRuns(outputDirectory);
        message += `, merged ${mergeCount} runs`;
      }
    }

    for (const xmlFile of xmlFiles) {
      escapeSmartQuotes(xmlFile);
    }

    return [null, message];
  } catch (e) {
    if (zipFailed) {
      return [null, `Error: ${inputFile} is not a valid Office file`];
    }
    return [null, `Error unpacking: ${e.message}`];
  }
}

function prettyPrintXml(xmlFile) {
  try {
    const content = fs.readFileSync(xmlFile, "utf-8");
    const dom = minidom.parseString(content);
    fs.writeFileSync(
      xmlFile,
      minidom.toprettyxml(dom, { indent: "  ", encoding: "utf-8" })
    );
  } catch (e) {
    // pass
  }
}

function escapeSmartQuotes(xmlFile) {
  try {
    let content = fs.readFileSync(xmlFile, "utf-8");
    for (const [char, entity] of Object.entries(SMART_QUOTE_REPLACEMENTS)) {
      content = content.split(char).join(entity);
    }
    fs.writeFileSync(xmlFile, content, { encoding: "utf-8" });
  } catch (e) {
    // pass
  }
}

function parseBool(x) {
  return String(x).toLowerCase() === "true";
}

function parseArgs(argv) {
  const positionals = [];
  const opts = { mergeRuns: true, simplifyRedlines: true };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--merge-runs") {
      opts.mergeRuns = parseBool(argv[++i]);
    } else if (a.startsWith("--merge-runs=")) {
      opts.mergeRuns = parseBool(a.split("=").slice(1).join("="));
    } else if (a === "--simplify-redlines") {
      opts.simplifyRedlines = parseBool(argv[++i]);
    } else if (a.startsWith("--simplify-redlines=")) {
      opts.simplifyRedlines = parseBool(a.split("=").slice(1).join("="));
    } else {
      positionals.push(a);
    }
  }
  return { positionals, opts };
}

function main() {
  const { positionals, opts } = parseArgs(process.argv.slice(2));
  if (positionals.length < 2) {
    process.stderr.write(
      "usage: unpack.js input_file output_directory [--merge-runs true|false] [--simplify-redlines true|false]\n"
    );
    process.exit(2);
  }

  const [inputFile, outputDirectory] = positionals;

  const [, message] = unpack(
    inputFile,
    outputDirectory,
    opts.mergeRuns,
    opts.simplifyRedlines
  );
  console.log(message);

  if (message.indexOf("Error") !== -1) {
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}

module.exports = { unpack };
