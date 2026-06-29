"use strict";
/*
 * Validator for PowerPoint presentation XML files against XSD schemas.
 * (Port of validators/pptx.py)
 */

const fs = require("fs");
const path = require("path");
const minidom = require("../../_pylib/minidom");
const pyfs = require("../../_pylib/pyfs");
const { BaseSchemaValidator } = require("./base");

const XMLNS_NS = "http://www.w3.org/2000/xmlns/";

function localOf(node) {
  return node.localName || node.tagName;
}
function lineOf(node) {
  return node.lineNumber != null ? node.lineNumber : undefined;
}
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
function findDescendantsClark(root, ns, local) {
  const out = [];
  for (const el of iterElements(root)) {
    if (el === root) continue;
    if (localOf(el) === local && el.namespaceURI === ns) out.push(el);
  }
  return out;
}
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
function getAttr(elem, name) {
  const v = elem.getAttribute(name);
  return v === "" && !elem.hasAttribute(name) ? null : v;
}
// Replicate Path.glob for "dir/.../<filepattern>" patterns, sorted.
function globRel(base, pattern) {
  const parts = pattern.split("/");
  const fileGlob = parts.pop();
  const dir = path.join(base, ...parts);
  let entries;
  try {
    entries = fs.readdirSync(dir);
  } catch (e) {
    return [];
  }
  const rx = new RegExp(
    "^" + fileGlob.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*") + "$"
  );
  return entries
    .filter((n) => rx.test(n))
    .sort()
    .map((n) => path.join(dir, n));
}
function stem(p) {
  const base = path.basename(p);
  const ext = path.extname(base);
  return ext ? base.slice(0, -ext.length) : base;
}

class PPTXSchemaValidator extends BaseSchemaValidator {
  async validate() {
    if (!this.validateXmlWellFormed()) return false;

    let allValid = true;
    if (!this.validateNamespaces()) allValid = false;
    if (!this.validateUniqueIds()) allValid = false;
    if (!this.validateUuidIds()) allValid = false;
    if (!this.validateFileReferences()) allValid = false;
    if (!this.validateSlideLayoutIds()) allValid = false;
    if (!this.validateContentTypes()) allValid = false;
    if (!(await this.validateAgainstXsd())) allValid = false;
    if (!this.validateNotesSlideReferences()) allValid = false;
    if (!this.validateAllRelationshipIds()) allValid = false;
    if (!this.validateNoDuplicateSlideLayouts()) allValid = false;

    return allValid;
  }

  validateUuidIds() {
    const errors = [];
    const uuidPattern = /^[\{\(]?[0-9A-Fa-f]{8}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{4}-?[0-9A-Fa-f]{12}[\}\)]?$/;

    for (const xmlFile of this.xml_files) {
      try {
        const root = this._parseOrThrow(xmlFile).documentElement;
        for (const elem of iterElements(root)) {
          for (const [attr, value] of attribItems(elem)) {
            const attrName = attr.split("}").pop().toLowerCase();
            if (attrName === "id" || attrName.endsWith("id")) {
              if (this._looksLikeUuid(value)) {
                if (!uuidPattern.test(value)) {
                  errors.push(
                    `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: ID '${value}' appears to be a UUID but contains invalid hex characters`
                  );
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
      console.log(`FAILED - Found ${errors.length} UUID ID validation errors:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All UUID-like IDs contain valid hex values");
    return true;
  }

  _looksLikeUuid(value) {
    const cleanValue = value.replace(/^[{}()]+/, "").replace(/[{}()]+$/, "").replace(/-/g, "");
    return (
      cleanValue.length === 32 &&
      [...cleanValue].every((c) => /[0-9A-Za-z]/.test(c))
    );
  }

  validateSlideLayoutIds() {
    const errors = [];

    const slideMasters = globRel(this.unpacked_dir, "ppt/slideMasters/*.xml");

    if (slideMasters.length === 0) {
      if (this.verbose) console.log("PASSED - No slide masters found");
      return true;
    }

    for (const slideMaster of slideMasters) {
      try {
        const root = this._parseOrThrow(slideMaster).documentElement;

        const relsFile = path.join(
          path.dirname(slideMaster),
          "_rels",
          `${path.basename(slideMaster)}.rels`
        );

        if (!pyfs.exists(relsFile)) {
          errors.push(
            `  ${this.rel(slideMaster)}: Missing relationships file: ${this.rel(relsFile)}`
          );
          continue;
        }

        const relsRoot = this._parseOrThrow(relsFile).documentElement;

        const validLayoutRids = new Set();
        for (const rel of findDescendantsClark(
          relsRoot,
          BaseSchemaValidator.PACKAGE_RELATIONSHIPS_NAMESPACE,
          "Relationship"
        )) {
          const relType = rel.getAttribute("Type") || "";
          if (relType.indexOf("slideLayout") !== -1) {
            validLayoutRids.add(rel.getAttribute("Id"));
          }
        }

        for (const sldLayoutId of findDescendantsClark(
          root,
          PPTXSchemaValidator.PRESENTATIONML_NAMESPACE,
          "sldLayoutId"
        )) {
          const rId = getNs(sldLayoutId, BaseSchemaValidator.OFFICE_RELATIONSHIPS_NAMESPACE, "id");
          const layoutId = getAttr(sldLayoutId, "id");

          if (rId && !validLayoutRids.has(rId)) {
            errors.push(
              `  ${this.rel(slideMaster)}: Line ${lineOf(sldLayoutId)}: sldLayoutId with id='${layoutId}' references r:id='${rId}' which is not found in slide layout relationships`
            );
          }
        }
      } catch (e) {
        errors.push(`  ${this.rel(slideMaster)}: Error: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} slide layout ID validation errors:`);
      for (const error of errors) console.log(error);
      console.log("Remove invalid references or add missing slide layouts to the relationships file.");
      return false;
    }
    if (this.verbose) console.log("PASSED - All slide layout IDs reference valid slide layouts");
    return true;
  }

  validateNoDuplicateSlideLayouts() {
    const errors = [];
    const slideRelsFiles = globRel(this.unpacked_dir, "ppt/slides/_rels/*.xml.rels");

    for (const relsFile of slideRelsFiles) {
      try {
        const root = this._parseOrThrow(relsFile).documentElement;

        const layoutRels = findDescendantsClark(
          root,
          BaseSchemaValidator.PACKAGE_RELATIONSHIPS_NAMESPACE,
          "Relationship"
        ).filter((rel) => (rel.getAttribute("Type") || "").indexOf("slideLayout") !== -1);

        if (layoutRels.length > 1) {
          errors.push(`  ${this.rel(relsFile)}: has ${layoutRels.length} slideLayout references`);
        }
      } catch (e) {
        errors.push(`  ${this.rel(relsFile)}: Error: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log("FAILED - Found slides with duplicate slideLayout references:");
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All slides have exactly one slideLayout reference");
    return true;
  }

  validateNotesSlideReferences() {
    const errors = [];
    const notesSlideReferences = {};

    const slideRelsFiles = globRel(this.unpacked_dir, "ppt/slides/_rels/*.xml.rels");

    if (slideRelsFiles.length === 0) {
      if (this.verbose) console.log("PASSED - No slide relationship files found");
      return true;
    }

    for (const relsFile of slideRelsFiles) {
      try {
        const root = this._parseOrThrow(relsFile).documentElement;

        for (const rel of findDescendantsClark(
          root,
          BaseSchemaValidator.PACKAGE_RELATIONSHIPS_NAMESPACE,
          "Relationship"
        )) {
          const relType = rel.getAttribute("Type") || "";
          if (relType.indexOf("notesSlide") !== -1) {
            const target = rel.getAttribute("Target") || "";
            if (target) {
              const normalizedTarget = target.split("../").join("");
              const slideName = stem(relsFile).split(".xml").join("");

              if (!notesSlideReferences[normalizedTarget]) {
                notesSlideReferences[normalizedTarget] = [];
              }
              notesSlideReferences[normalizedTarget].push([slideName, relsFile]);
            }
          }
        }
      } catch (e) {
        errors.push(`  ${this.rel(relsFile)}: Error: ${e.message}`);
      }
    }

    for (const target of Object.keys(notesSlideReferences)) {
      const references = notesSlideReferences[target];
      if (references.length > 1) {
        const slideNames = references.map((ref) => ref[0]);
        errors.push(
          `  Notes slide '${target}' is referenced by multiple slides: ${slideNames.join(", ")}`
        );
        for (const [, relsFile] of references) {
          errors.push(`    - ${this.rel(relsFile)}`);
        }
      }
    }

    if (errors.length) {
      const headCount = errors.filter((e) => !e.startsWith("    ")).length;
      console.log(`FAILED - Found ${headCount} notes slide reference validation errors:`);
      for (const error of errors) console.log(error);
      console.log("Each slide may optionally have its own slide file.");
      return false;
    }
    if (this.verbose) console.log("PASSED - All notes slide references are unique");
    return true;
  }
}

PPTXSchemaValidator.PRESENTATIONML_NAMESPACE = "http://schemas.openxmlformats.org/presentationml/2006/main";

PPTXSchemaValidator.ELEMENT_RELATIONSHIP_TYPES = {
  sldid: "slide",
  sldmasterid: "slidemaster",
  notesmasterid: "notesmaster",
  sldlayoutid: "slidelayout",
  themeid: "theme",
  tablestyleid: "tablestyles",
};

module.exports = { PPTXSchemaValidator };
