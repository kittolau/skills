"use strict";
/*
 * Base validator with common validation logic for document files.
 * (Port of validators/base.py)
 *
 * NOTE on XSD validation: the Python original uses lxml.etree.XMLSchema.
 * A native lxml/libxml build is not available in this environment, so XSD
 * validation is performed with the pure-WASM `xmllint-wasm` (libxml2). The
 * pass/fail decision logic is identical (set-difference of edited-vs-original
 * errors, with IGNORED patterns filtered); only the underlying engine's raw
 * error-message strings differ from lxml. Methods that touch XSD are async.
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const AdmZip = require("adm-zip");
const minidom = require("../../_pylib/minidom");
const pyfs = require("../../_pylib/pyfs");

const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

function clarkTag(node) {
  const ns = node.namespaceURI;
  const local = node.localName || node.tagName;
  return ns ? `{${ns}}${local}` : local;
}

function localOf(node) {
  return node.localName || node.tagName;
}

function lineOf(node) {
  return node.lineNumber != null ? node.lineNumber : undefined;
}

// Iterate element node + all descendant elements, document order (like root.iter()).
function iterElements(root) {
  const out = [];
  (function walk(n) {
    if (n.nodeType === minidom.ELEMENT_NODE) {
      out.push(n);
      for (const c of minidom.childNodesArray(n)) walk(c);
    }
  })(root);
  return out;
}

function iterAncestors(node) {
  const out = [];
  let p = node.parentNode;
  while (p && p.nodeType === minidom.ELEMENT_NODE) {
    out.push(p);
    p = p.parentNode;
  }
  return out;
}

// Descendant elements (excludes root) matching ns+local — mirrors ".//prefix:local".
function findDescendantsNs(root, ns, local) {
  const out = [];
  for (const el of iterElements(root)) {
    if (el === root) continue;
    if (localOf(el) === local && el.namespaceURI === ns) out.push(el);
  }
  return out;
}

function findDescendantsClark(root, ns, local) {
  // Includes root if it matches? lxml findall(".//{ns}local") excludes root.
  return findDescendantsNs(root, ns, local);
}

// elem.attrib.items() equivalent: Clark-notation keys, EXCLUDING xmlns decls.
function attribItems(elem) {
  const out = [];
  const attrs = elem.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a.namespaceURI === XMLNS_NS) continue;
    if (a.name === "xmlns" || a.name.startsWith("xmlns:")) continue;
    const key = a.namespaceURI ? `{${a.namespaceURI}}${a.localName}` : a.name;
    out.push([key, a.value]);
  }
  return out;
}

function getNs(elem, ns, local) {
  if (elem.getAttributeNS) {
    const v = elem.getAttributeNS(ns, local);
    if (v) return v;
  }
  return null;
}

function nsmapPrefixes(root) {
  // Prefixes declared on root (lxml root.nsmap keys minus default None).
  const prefixes = new Set();
  const attrs = root.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a.name === "xmlns") continue; // default namespace -> None key
    if (a.name.startsWith("xmlns:")) prefixes.add(a.name.slice(6));
  }
  return prefixes;
}

function pyRepr(s) {
  // Approximate Python repr() of a short string (single quotes).
  let r = "'";
  for (const ch of s) {
    if (ch === "\\") r += "\\\\";
    else if (ch === "'") r += "\\'";
    else if (ch === "\n") r += "\\n";
    else if (ch === "\t") r += "\\t";
    else if (ch === "\r") r += "\\r";
    else r += ch;
  }
  return r + "'";
}

class BaseSchemaValidator {
  constructor(unpackedDir, originalFile, verbose) {
    this.unpacked_dir = path.resolve(unpackedDir);
    this.original_file = originalFile ? originalFile : null;
    this.verbose = !!verbose;

    this.schemas_dir = path.join(__dirname, "..", "schemas");

    this.xml_files = pyfs
      .rglob(this.unpacked_dir, "*.xml")
      .concat(pyfs.rglob(this.unpacked_dir, "*.rels"));

    if (this.xml_files.length === 0) {
      console.log(`Warning: No XML files found in ${this.unpacked_dir}`);
    }

    this._schemaFilesCache = null;
  }

  rel(p) {
    return pyfs.relativeTo(this.unpacked_dir, p);
  }

  validate() {
    throw new Error("Subclasses must implement the validate method");
  }

  repair() {
    return this.repairWhitespacePreservation();
  }

  repairWhitespacePreservation() {
    let repairs = 0;
    for (const xmlFile of this.xml_files) {
      try {
        const content = fs.readFileSync(xmlFile, "utf-8");
        const dom = minidom.parseString(content);
        let modified = false;

        const all = dom.getElementsByTagName("*");
        for (let i = 0; i < all.length; i++) {
          const elem = all[i];
          if (elem.tagName.endsWith(":t") && elem.firstChild) {
            const text = elem.firstChild.nodeValue;
            if (
              text &&
              (text.startsWith(" ") ||
                text.startsWith("\t") ||
                text.endsWith(" ") ||
                text.endsWith("\t"))
            ) {
              if (elem.getAttribute("xml:space") !== "preserve") {
                elem.setAttribute("xml:space", "preserve");
                const textPreview =
                  text.length > 30
                    ? pyRepr(text.slice(0, 30)) + "..."
                    : pyRepr(text);
                console.log(
                  `  Repaired: ${path.basename(xmlFile)}: Added xml:space='preserve' to ${elem.tagName}: ${textPreview}`
                );
                repairs += 1;
                modified = true;
              }
            }
          }
        }

        if (modified) {
          fs.writeFileSync(xmlFile, minidom.toxml(dom, "UTF-8"));
        }
      } catch (e) {
        // pass
      }
    }
    return repairs;
  }

  _parseOrThrow(xmlFile) {
    return minidom.parseString(fs.readFileSync(xmlFile, "utf-8"));
  }

  validateXmlWellFormed() {
    const errors = [];
    for (const xmlFile of this.xml_files) {
      try {
        this._parseOrThrow(xmlFile);
      } catch (e) {
        errors.push(`  ${this.rel(xmlFile)}: Line ${e.lineno || 0}: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} XML violations:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All XML files are well-formed");
    return true;
  }

  validateNamespaces() {
    const errors = [];
    for (const xmlFile of this.xml_files) {
      let root;
      try {
        root = this._parseOrThrow(xmlFile).documentElement;
      } catch (e) {
        continue;
      }
      const declared = nsmapPrefixes(root);
      for (const [k, v] of attribItems(root)) {
        if (k.endsWith("Ignorable")) {
          const used = v.split(/\s+/).filter((x) => x.length > 0);
          for (const ns of used) {
            if (!declared.has(ns)) {
              errors.push(
                `  ${this.rel(xmlFile)}: Namespace '${ns}' in Ignorable but not declared`
              );
            }
          }
        }
      }
    }

    if (errors.length) {
      console.log(`FAILED - ${errors.length} namespace issues:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All namespace prefixes properly declared");
    return true;
  }

  validateUniqueIds() {
    const errors = [];
    const globalIds = {};

    for (const xmlFile of this.xml_files) {
      try {
        const root = this._parseOrThrow(xmlFile).documentElement;
        const fileIds = {};

        const mcElements = findDescendantsNs(
          root,
          BaseSchemaValidator.MC_NAMESPACE,
          "AlternateContent"
        );
        for (const elem of mcElements) {
          if (elem.parentNode) elem.parentNode.removeChild(elem);
        }

        for (const elem of iterElements(root)) {
          const tag = localOf(elem).toLowerCase();

          if (Object.prototype.hasOwnProperty.call(BaseSchemaValidator.UNIQUE_ID_REQUIREMENTS, tag)) {
            const inExcluded = iterAncestors(elem).some((anc) =>
              BaseSchemaValidator.EXCLUDED_ID_CONTAINERS.has(localOf(anc).toLowerCase())
            );
            if (inExcluded) continue;

            const [attrName, scope] = BaseSchemaValidator.UNIQUE_ID_REQUIREMENTS[tag];

            let idValue = null;
            for (const [attr, value] of attribItems(elem)) {
              const attrLocal = attr.indexOf("}") !== -1
                ? attr.split("}").pop().toLowerCase()
                : attr.toLowerCase();
              if (attrLocal === attrName) {
                idValue = value;
                break;
              }
            }

            if (idValue !== null) {
              if (scope === "global") {
                if (Object.prototype.hasOwnProperty.call(globalIds, idValue)) {
                  const [prevFile, prevLine, prevTag] = globalIds[idValue];
                  errors.push(
                    `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: Global ID '${idValue}' in <${tag}> already used in ${prevFile} at line ${prevLine} in <${prevTag}>`
                  );
                } else {
                  globalIds[idValue] = [this.rel(xmlFile), lineOf(elem), tag];
                }
              } else if (scope === "file") {
                const key = `${tag}\u0000${attrName}`;
                if (!fileIds[key]) fileIds[key] = {};
                if (Object.prototype.hasOwnProperty.call(fileIds[key], idValue)) {
                  const prevLine = fileIds[key][idValue];
                  errors.push(
                    `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: Duplicate ${attrName}='${idValue}' in <${tag}> (first occurrence at line ${prevLine})`
                  );
                } else {
                  fileIds[key][idValue] = lineOf(elem);
                }
              }
            }
          }
        }
      } catch (e) {
        errors.push(`  ${this.rel(xmlFile)}: Error: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} ID uniqueness violations:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All required IDs are unique");
    return true;
  }

  validateFileReferences() {
    const errors = [];

    const relsFiles = pyfs.rglob(this.unpacked_dir, "*.rels");

    if (relsFiles.length === 0) {
      if (this.verbose) console.log("PASSED - No .rels files found");
      return true;
    }

    const allFiles = [];
    for (const filePath of pyfs.rglob(this.unpacked_dir, "*")) {
      const base = path.basename(filePath);
      if (
        pyfs.isFile(filePath) &&
        base !== "[Content_Types].xml" &&
        !base.endsWith(".rels")
      ) {
        allFiles.push(path.resolve(filePath));
      }
    }

    const allReferenced = new Set();

    if (this.verbose) {
      console.log(`Found ${relsFiles.length} .rels files and ${allFiles.length} target files`);
    }

    for (const relsFile of relsFiles) {
      try {
        const relsRoot = this._parseOrThrow(relsFile).documentElement;
        const relsDir = path.dirname(relsFile);
        const brokenRefs = [];

        for (const rel of findDescendantsClark(
          relsRoot,
          BaseSchemaValidator.PACKAGE_RELATIONSHIPS_NAMESPACE,
          "Relationship"
        )) {
          const target = rel.getAttribute("Target");
          if (target && !(target.startsWith("http") || target.startsWith("mailto:"))) {
            let targetPath;
            if (target.startsWith("/")) {
              targetPath = path.join(this.unpacked_dir, target.replace(/^\/+/, ""));
            } else if (path.basename(relsFile) === ".rels") {
              targetPath = path.join(this.unpacked_dir, target);
            } else {
              const baseDir = path.dirname(relsDir);
              targetPath = path.join(baseDir, target);
            }

            try {
              targetPath = path.resolve(targetPath);
              if (pyfs.exists(targetPath) && pyfs.isFile(targetPath)) {
                allReferenced.add(targetPath);
              } else {
                brokenRefs.push([target, lineOf(rel)]);
              }
            } catch (e) {
              brokenRefs.push([target, lineOf(rel)]);
            }
          }
        }

        if (brokenRefs.length) {
          const relPath = this.rel(relsFile);
          for (const [brokenRef, lineNum] of brokenRefs) {
            errors.push(`  ${relPath}: Line ${lineNum}: Broken reference to ${brokenRef}`);
          }
        }
      } catch (e) {
        errors.push(`  Error parsing ${this.rel(relsFile)}: ${e.message}`);
      }
    }

    const unreferenced = allFiles.filter((f) => !allReferenced.has(f));
    if (unreferenced.length) {
      unreferenced.sort();
      for (const unrefFile of unreferenced) {
        errors.push(`  Unreferenced file: ${this.rel(unrefFile)}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} relationship validation errors:`);
      for (const error of errors) console.log(error);
      console.log(
        "CRITICAL: These errors will cause the document to appear corrupt. " +
          "Broken references MUST be fixed, " +
          "and unreferenced files MUST be referenced or removed."
      );
      return false;
    }
    if (this.verbose) {
      console.log("PASSED - All references are valid and all files are properly referenced");
    }
    return true;
  }

  validateAllRelationshipIds() {
    const errors = [];

    for (const xmlFile of this.xml_files) {
      if (path.extname(xmlFile) === ".rels") continue;

      const relsDir = path.join(path.dirname(xmlFile), "_rels");
      const relsFile = path.join(relsDir, `${path.basename(xmlFile)}.rels`);

      if (!pyfs.exists(relsFile)) continue;

      try {
        const relsRoot = this._parseOrThrow(relsFile).documentElement;
        const ridToType = {};

        for (const rel of findDescendantsClark(
          relsRoot,
          BaseSchemaValidator.PACKAGE_RELATIONSHIPS_NAMESPACE,
          "Relationship"
        )) {
          const rid = rel.getAttribute("Id");
          const relType = rel.getAttribute("Type") || "";
          if (rid) {
            if (Object.prototype.hasOwnProperty.call(ridToType, rid)) {
              errors.push(
                `  ${this.rel(relsFile)}: Line ${lineOf(rel)}: Duplicate relationship ID '${rid}' (IDs must be unique)`
              );
            }
            const typeName = relType.indexOf("/") !== -1 ? relType.split("/").pop() : relType;
            ridToType[rid] = typeName;
          }
        }

        const xmlRoot = this._parseOrThrow(xmlFile).documentElement;
        const rNs = BaseSchemaValidator.OFFICE_RELATIONSHIPS_NAMESPACE;
        const ridAttrs = ["id", "embed", "link"];

        for (const elem of iterElements(xmlRoot)) {
          for (const attrName of ridAttrs) {
            const ridAttr = getNs(elem, rNs, attrName);
            if (!ridAttr) continue;
            const elemName = localOf(elem);

            if (!Object.prototype.hasOwnProperty.call(ridToType, ridAttr)) {
              const keys = Object.keys(ridToType).sort();
              const preview = keys.slice(0, 5).join(", ") + (keys.length > 5 ? "..." : "");
              errors.push(
                `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: <${elemName}> r:${attrName} references non-existent relationship '${ridAttr}' (valid IDs: ${preview})`
              );
            } else if (attrName === "id" && Object.keys(this.constructor.ELEMENT_RELATIONSHIP_TYPES || {}).length) {
              const expectedType = this._getExpectedRelationshipType(elemName);
              if (expectedType) {
                const actualType = ridToType[ridAttr];
                if (actualType.toLowerCase().indexOf(expectedType) === -1) {
                  errors.push(
                    `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: <${elemName}> references '${ridAttr}' which points to '${actualType}' but should point to a '${expectedType}' relationship`
                  );
                }
              }
            }
          }
        }
      } catch (e) {
        errors.push(`  Error processing ${this.rel(xmlFile)}: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} relationship ID reference errors:`);
      for (const error of errors) console.log(error);
      console.log("\nThese ID mismatches will cause the document to appear corrupt!");
      return false;
    }
    if (this.verbose) console.log("PASSED - All relationship ID references are valid");
    return true;
  }

  _getExpectedRelationshipType(elementName) {
    const elemLower = elementName.toLowerCase();
    const map = this.constructor.ELEMENT_RELATIONSHIP_TYPES || {};

    if (Object.prototype.hasOwnProperty.call(map, elemLower)) return map[elemLower];

    if (elemLower.endsWith("id") && elemLower.length > 2) {
      const prefix = elemLower.slice(0, -2);
      if (prefix.endsWith("master")) return prefix;
      if (prefix.endsWith("layout")) return prefix;
      if (prefix === "sld") return "slide";
      return prefix;
    }

    if (elemLower.endsWith("reference") && elemLower.length > 9) {
      return elemLower.slice(0, -9);
    }

    return null;
  }

  validateContentTypes() {
    const errors = [];

    const contentTypesFile = path.join(this.unpacked_dir, "[Content_Types].xml");
    if (!pyfs.exists(contentTypesFile)) {
      console.log("FAILED - [Content_Types].xml file not found");
      return false;
    }

    try {
      const root = this._parseOrThrow(contentTypesFile).documentElement;
      const declaredParts = new Set();
      const declaredExtensions = new Set();

      for (const override of findDescendantsClark(
        root,
        BaseSchemaValidator.CONTENT_TYPES_NAMESPACE,
        "Override"
      )) {
        const partName = override.getAttribute("PartName");
        if (partName) declaredParts.add(partName.replace(/^\/+/, ""));
      }

      for (const def of findDescendantsClark(
        root,
        BaseSchemaValidator.CONTENT_TYPES_NAMESPACE,
        "Default"
      )) {
        const extension = def.getAttribute("Extension");
        if (extension) declaredExtensions.add(extension.toLowerCase());
      }

      const declarableRoots = new Set([
        "sld", "sldLayout", "sldMaster", "presentation",
        "document", "workbook", "worksheet", "theme",
      ]);

      const mediaExtensions = {
        png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg",
        gif: "image/gif", bmp: "image/bmp", tiff: "image/tiff",
        wmf: "image/x-wmf", emf: "image/x-emf",
      };

      const allFiles = pyfs.rglob(this.unpacked_dir, "*").filter((f) => pyfs.isFile(f));

      for (const xmlFile of this.xml_files) {
        const pathStr = this.rel(xmlFile).split(path.sep).join("/");
        if ([".rels", "[Content_Types]", "docProps/", "_rels/"].some((s) => pathStr.indexOf(s) !== -1)) {
          continue;
        }
        try {
          const rootTag = clarkTag(this._parseOrThrow(xmlFile).documentElement);
          const rootName = rootTag.indexOf("}") !== -1 ? rootTag.split("}").pop() : rootTag;
          if (declarableRoots.has(rootName) && !declaredParts.has(pathStr)) {
            errors.push(`  ${pathStr}: File with <${rootName}> root not declared in [Content_Types].xml`);
          }
        } catch (e) {
          continue;
        }
      }

      for (const filePath of allFiles) {
        const ext = path.extname(filePath).toLowerCase();
        if (ext === ".xml" || ext === ".rels") continue;
        if (path.basename(filePath) === "[Content_Types].xml") continue;
        const parts = this.rel(filePath).split(path.sep);
        if (parts.includes("_rels") || parts.includes("docProps")) continue;

        const extension = path.extname(filePath).replace(/^\./, "").toLowerCase();
        if (extension && !declaredExtensions.has(extension)) {
          if (Object.prototype.hasOwnProperty.call(mediaExtensions, extension)) {
            const relativePath = this.rel(filePath);
            errors.push(
              `  ${relativePath}: File with extension '${extension}' not declared in [Content_Types].xml - should add: <Default Extension="${extension}" ContentType="${mediaExtensions[extension]}"/>`
            );
          }
        }
      }
    } catch (e) {
      errors.push(`  Error parsing [Content_Types].xml: ${e.message}`);
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} content type declaration errors:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) {
      console.log("PASSED - All content files are properly declared in [Content_Types].xml");
    }
    return true;
  }

  // ----------------------- XSD validation (async) -----------------------

  async validateFileAgainstXsd(xmlFile, verbose) {
    xmlFile = path.resolve(xmlFile);
    const unpackedDir = path.resolve(this.unpacked_dir);

    const [isValid, currentErrors] = await this._validateSingleFileXsd(xmlFile, unpackedDir);

    if (isValid === null) return [null, new Set()];
    if (isValid) return [true, new Set()];

    const originalErrors = await this._getOriginalFileErrors(xmlFile);

    let newErrors = new Set([...currentErrors].filter((e) => !originalErrors.has(e)));
    newErrors = new Set(
      [...newErrors].filter(
        (e) => !BaseSchemaValidator.IGNORED_VALIDATION_ERRORS.some((p) => e.indexOf(p) !== -1)
      )
    );

    if (newErrors.size) {
      if (verbose) {
        const relativePath = pyfs.relativeTo(unpackedDir, xmlFile);
        console.log(`FAILED - ${relativePath}: ${newErrors.size} new error(s)`);
        for (const error of [...newErrors].slice(0, 3)) {
          const truncated = error.length > 250 ? error.slice(0, 250) + "..." : error;
          console.log(`  - ${truncated}`);
        }
      }
      return [false, newErrors];
    }
    if (verbose) {
      console.log(`PASSED - No new errors (original had ${currentErrors.size} errors)`);
    }
    return [true, new Set()];
  }

  async validateAgainstXsd() {
    const newErrors = [];
    let originalErrorCount = 0;
    let validCount = 0;
    let skippedCount = 0;

    for (const xmlFile of this.xml_files) {
      const relativePath = this.rel(xmlFile);
      const [isValid, newFileErrors] = await this.validateFileAgainstXsd(xmlFile, false);

      if (isValid === null) {
        skippedCount += 1;
        continue;
      } else if (isValid && newFileErrors.size === 0) {
        validCount += 1;
        continue;
      } else if (isValid) {
        originalErrorCount += 1;
        validCount += 1;
        continue;
      }

      newErrors.push(`  ${relativePath}: ${newFileErrors.size} new error(s)`);
      for (const error of [...newFileErrors].slice(0, 3)) {
        newErrors.push(error.length > 250 ? `    - ${error.slice(0, 250)}...` : `    - ${error}`);
      }
    }

    if (this.verbose) {
      console.log(`Validated ${this.xml_files.length} files:`);
      console.log(`  - Valid: ${validCount}`);
      console.log(`  - Skipped (no schema): ${skippedCount}`);
      if (originalErrorCount) {
        console.log(`  - With original errors (ignored): ${originalErrorCount}`);
      }
      const newCount = newErrors.length > 0
        ? newErrors.filter((e) => !e.startsWith("    ")).length
        : 0;
      console.log(`  - With NEW errors: ${newCount}`);
    }

    if (newErrors.length) {
      console.log("\nFAILED - Found NEW validation errors:");
      for (const error of newErrors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("\nPASSED - No new XSD validation errors introduced");
    return true;
  }

  _getSchemaPath(xmlFile) {
    const name = path.basename(xmlFile);
    const M = BaseSchemaValidator.SCHEMA_MAPPINGS;

    if (Object.prototype.hasOwnProperty.call(M, name)) {
      return path.join(this.schemas_dir, M[name]);
    }
    if (path.extname(xmlFile) === ".rels") {
      return path.join(this.schemas_dir, M[".rels"]);
    }
    const sx = xmlFile.split(path.sep).join("/");
    if (sx.indexOf("charts/") !== -1 && name.startsWith("chart")) {
      return path.join(this.schemas_dir, M["chart"]);
    }
    if (sx.indexOf("theme/") !== -1 && name.startsWith("theme")) {
      return path.join(this.schemas_dir, M["theme"]);
    }
    const parentName = path.basename(path.dirname(xmlFile));
    if (BaseSchemaValidator.MAIN_CONTENT_FOLDERS.has(parentName)) {
      return path.join(this.schemas_dir, M[parentName]);
    }
    return null;
  }

  _loadSchemaFiles() {
    if (this._schemaFilesCache) return this._schemaFilesCache;
    const files = [];
    // The OOXML schemas cross-reference each other by basename only (e.g.
    // <xs:include schemaLocation="dml-main.xsd"/>), and xmllint-wasm mounts
    // every file into a single flat virtual FS (nested directories trigger an
    // "ErrnoError: FS error"). All 39 schema basenames are unique, so we key
    // the virtual files by basename to satisfy both constraints.
    for (const f of pyfs.rglob(this.schemas_dir, "*.xsd")) {
      const fileName = path.basename(f);
      files.push({ fileName, contents: fs.readFileSync(f, "utf-8") });
    }
    this._schemaFilesCache = files;
    return files;
  }


  // Mutating DOM helpers replicating lxml preprocessing on a fresh parse.
  _removeIgnorableElements(node) {
    const toRemove = [];
    for (const child of minidom.childNodesArray(node)) {
      if (child.nodeType !== minidom.ELEMENT_NODE) continue;
      const ns = child.namespaceURI;
      if (ns && !BaseSchemaValidator.OOXML_NAMESPACES.has(ns)) {
        toRemove.push(child);
        continue;
      }
      this._removeIgnorableElements(child);
    }
    for (const el of toRemove) node.removeChild(el);
  }

  _cleanIgnorableNamespaces(root) {
    for (const elem of iterElements(root)) {
      const attrs = elem.attributes;
      const remove = [];
      for (let i = 0; i < attrs.length; i++) {
        const a = attrs.item(i);
        if (a.namespaceURI === XMLNS_NS || a.name === "xmlns" || a.name.startsWith("xmlns:")) {
          continue;
        }
        if (a.namespaceURI && !BaseSchemaValidator.OOXML_NAMESPACES.has(a.namespaceURI)) {
          remove.push(a);
        }
      }
      for (const a of remove) elem.removeAttributeNode(a);
    }
    this._removeIgnorableElements(root);
  }

  _preprocessForMcIgnorable(root) {
    if (root.getAttributeNS) {
      if (root.getAttributeNodeNS) {
        const node = root.getAttributeNodeNS(BaseSchemaValidator.MC_NAMESPACE, "Ignorable");
        if (node) root.removeAttributeNode(node);
      }
    }
  }

  _removeTemplateTagsFromTextNodes(root) {
    const pattern = /\{\{[^}]*\}\}/g;
    for (const elem of iterElements(root)) {
      const tag = clarkTag(elem);
      if (tag.endsWith("}t") || tag === "t") continue;

      // .text  -> leading text-node child
      const first = elem.firstChild;
      if (first && first.nodeType === minidom.TEXT_NODE) {
        first.data = first.data.replace(pattern, "");
      }
      // .tail  -> following text-node sibling
      const next = elem.nextSibling;
      if (next && next.nodeType === minidom.TEXT_NODE) {
        next.data = next.data.replace(pattern, "");
      }
    }
  }

  async _validateSingleFileXsd(xmlFile, basePath) {
    const schemaPath = this._getSchemaPath(xmlFile);
    if (!schemaPath) return [null, null];

    try {
      const { validateXML } = require("xmllint-wasm");

      const dom = this._parseOrThrow(xmlFile);
      const root = dom.documentElement;

      this._removeTemplateTagsFromTextNodes(root);
      this._preprocessForMcIgnorable(root);

      const relParts = pyfs.relativeTo(basePath, xmlFile).split(path.sep);
      if (relParts.length && BaseSchemaValidator.MAIN_CONTENT_FOLDERS.has(relParts[0])) {
        this._cleanIgnorableNamespaces(root);
      }

      const xmlString = minidom.toxml(dom, "UTF-8");

      const schemaFiles = this._loadSchemaFiles();
      // Schemas are mounted by basename (see _loadSchemaFiles).
      const schemaRel = path.basename(schemaPath);
      const schemaEntry = schemaFiles.find((s) => s.fileName === schemaRel);
      const preload = schemaFiles.filter((s) => s.fileName !== schemaRel);


      const result = await validateXML({
        xml: [{ fileName: path.basename(xmlFile), contents: xmlString }],
        schema: [schemaEntry],
        preload: preload,
        initialMemoryPages: 256,
        maxMemoryPages: 4096,
      });

      if (result.valid) return [true, new Set()];

      const errors = new Set();
      for (const e of result.errors || []) {
        errors.add(e.message != null ? e.message : String(e));
      }
      return [false, errors];
    } catch (e) {
      return [false, new Set([String(e.message != null ? e.message : e)])];
    }
  }

  async _getOriginalFileErrors(xmlFile) {
    if (this.original_file === null) return new Set();

    xmlFile = path.resolve(xmlFile);
    const unpackedDir = path.resolve(this.unpacked_dir);
    const relativePath = pyfs.relativeTo(unpackedDir, xmlFile);

    const tempPath = fs.mkdtempSync(path.join(os.tmpdir(), "ooxml-orig-"));
    try {
      const zip = new AdmZip(this.original_file);
      zip.extractAllTo(tempPath, true);

      const originalXmlFile = path.join(tempPath, relativePath);
      if (!pyfs.exists(originalXmlFile)) return new Set();

      const [, errors] = await this._validateSingleFileXsd(originalXmlFile, tempPath);
      return errors ? errors : new Set();
    } finally {
      try {
        fs.rmSync(tempPath, { recursive: true, force: true });
      } catch (e) {}
    }
  }
}

BaseSchemaValidator.IGNORED_VALIDATION_ERRORS = [
  "hyphenationZone",
  "purl.org/dc/terms",
];

BaseSchemaValidator.UNIQUE_ID_REQUIREMENTS = {
  comment: ["id", "file"],
  commentrangestart: ["id", "file"],
  commentrangeend: ["id", "file"],
  bookmarkstart: ["id", "file"],
  bookmarkend: ["id", "file"],
  sldid: ["id", "file"],
  sldmasterid: ["id", "global"],
  sldlayoutid: ["id", "global"],
  cm: ["authorid", "file"],
  sheet: ["sheetid", "file"],
  definedname: ["id", "file"],
  cxnsp: ["id", "file"],
  sp: ["id", "file"],
  pic: ["id", "file"],
  grpsp: ["id", "file"],
};

BaseSchemaValidator.EXCLUDED_ID_CONTAINERS = new Set(["sectionlst"]);

BaseSchemaValidator.ELEMENT_RELATIONSHIP_TYPES = {};

BaseSchemaValidator.SCHEMA_MAPPINGS = {
  word: "ISO-IEC29500-4_2016/wml.xsd",
  ppt: "ISO-IEC29500-4_2016/pml.xsd",
  xl: "ISO-IEC29500-4_2016/sml.xsd",
  "[Content_Types].xml": "ecma/fouth-edition/opc-contentTypes.xsd",
  "app.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesExtended.xsd",
  "core.xml": "ecma/fouth-edition/opc-coreProperties.xsd",
  "custom.xml": "ISO-IEC29500-4_2016/shared-documentPropertiesCustom.xsd",
  ".rels": "ecma/fouth-edition/opc-relationships.xsd",
  "people.xml": "microsoft/wml-2012.xsd",
  "commentsIds.xml": "microsoft/wml-cid-2016.xsd",
  "commentsExtensible.xml": "microsoft/wml-cex-2018.xsd",
  "commentsExtended.xml": "microsoft/wml-2012.xsd",
  chart: "ISO-IEC29500-4_2016/dml-chart.xsd",
  theme: "ISO-IEC29500-4_2016/dml-main.xsd",
  drawing: "ISO-IEC29500-4_2016/dml-main.xsd",
};

BaseSchemaValidator.MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
BaseSchemaValidator.XML_NAMESPACE = "http://www.w3.org/XML/1998/namespace";
BaseSchemaValidator.PACKAGE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships";
BaseSchemaValidator.OFFICE_RELATIONSHIPS_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
BaseSchemaValidator.CONTENT_TYPES_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/content-types";
BaseSchemaValidator.MAIN_CONTENT_FOLDERS = new Set(["word", "ppt", "xl"]);

BaseSchemaValidator.OOXML_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/math",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://schemas.openxmlformats.org/schemaLibrary/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "http://schemas.openxmlformats.org/drawingml/2006/chartDrawing",
  "http://schemas.openxmlformats.org/drawingml/2006/diagram",
  "http://schemas.openxmlformats.org/drawingml/2006/picture",
  "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://schemas.openxmlformats.org/presentationml/2006/main",
  "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  "http://schemas.openxmlformats.org/officeDocument/2006/sharedTypes",
  "http://www.w3.org/XML/1998/namespace",
]);

module.exports = { BaseSchemaValidator };
