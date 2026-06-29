"use strict";
/*
 * Validator for Word document XML files against XSD schemas.
 * (Port of validators/docx.py)
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const AdmZip = require("adm-zip");
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
function matchNs(el, ns, local) {
  return localOf(el) === local && el.namespaceURI === ns;
}
function iterByTag(root, ns, local) {
  return iterElements(root).filter((el) => matchNs(el, ns, local));
}
function findDescendants(root, ns, local) {
  return iterElements(root).filter((el) => el !== root && matchNs(el, ns, local));
}
function hasAncestor(el, ns, local) {
  let p = el.parentNode;
  while (p && p.nodeType === minidom.ELEMENT_NODE) {
    if (matchNs(p, ns, local)) return true;
    p = p.parentNode;
  }
  return false;
}
function getNs(elem, ns, local) {
  if (elem.getAttributeNS) {
    const v = elem.getAttributeNS(ns, local);
    if (v !== "" || (elem.hasAttributeNS && elem.hasAttributeNS(ns, local))) {
      return v === "" ? "" : v;
    }
  }
  return null;
}
// elem text == lxml .text: leading text-node child value (or null)
function elemText(el) {
  const first = el.firstChild;
  if (first && first.nodeType === minidom.TEXT_NODE) return first.data;
  return null;
}
function pyReprStr(s) {
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

class DOCXSchemaValidator extends BaseSchemaValidator {
  async validate() {
    if (!this.validateXmlWellFormed()) return false;

    let allValid = true;
    if (!this.validateNamespaces()) allValid = false;
    if (!this.validateUniqueIds()) allValid = false;
    if (!this.validateFileReferences()) allValid = false;
    if (!this.validateContentTypes()) allValid = false;
    if (!(await this.validateAgainstXsd())) allValid = false;
    if (!this.validateWhitespacePreservation()) allValid = false;
    if (!this.validateDeletions()) allValid = false;
    if (!this.validateInsertions()) allValid = false;
    if (!this.validateAllRelationshipIds()) allValid = false;
    if (!this.validateIdConstraints()) allValid = false;
    if (!this.validateCommentMarkers()) allValid = false;

    this.compareParagraphCounts();

    return allValid;
  }

  validateWhitespacePreservation() {
    const errors = [];

    for (const xmlFile of this.xml_files) {
      if (path.basename(xmlFile) !== "document.xml") continue;

      try {
        const root = this._parseOrThrow(xmlFile).documentElement;
        for (const elem of iterByTag(root, DOCXSchemaValidator.WORD_2006_NAMESPACE, "t")) {
          const text = elemText(elem);
          if (text) {
            if (/^[ \t\n\r]/.test(text) || /[ \t\n\r]$/.test(text)) {
              const xmlSpace = getNs(elem, BaseSchemaValidator.XML_NAMESPACE, "space");
              if (xmlSpace === null || xmlSpace !== "preserve") {
                const rep = pyReprStr(text);
                const textPreview = rep.length > 50 ? rep.slice(0, 50) + "..." : rep;
                errors.push(
                  `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: w:t element with whitespace missing xml:space='preserve': ${textPreview}`
                );
              }
            }
          }
        }
      } catch (e) {
        errors.push(`  ${this.rel(xmlFile)}: Error: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} whitespace preservation violations:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All whitespace is properly preserved");
    return true;
  }

  validateDeletions() {
    const errors = [];
    const W = DOCXSchemaValidator.WORD_2006_NAMESPACE;

    for (const xmlFile of this.xml_files) {
      if (path.basename(xmlFile) !== "document.xml") continue;

      try {
        const root = this._parseOrThrow(xmlFile).documentElement;

        for (const tElem of iterElements(root)) {
          if (!matchNs(tElem, W, "t")) continue;
          if (!hasAncestor(tElem, W, "del")) continue;
          const txt = elemText(tElem);
          if (txt) {
            const rep = pyReprStr(txt);
            const textPreview = rep.length > 50 ? rep.slice(0, 50) + "..." : rep;
            errors.push(
              `  ${this.rel(xmlFile)}: Line ${lineOf(tElem)}: <w:t> found within <w:del>: ${textPreview}`
            );
          }
        }

        for (const instrElem of iterElements(root)) {
          if (!matchNs(instrElem, W, "instrText")) continue;
          if (!hasAncestor(instrElem, W, "del")) continue;
          const t = elemText(instrElem) || "";
          const rep = pyReprStr(t);
          const textPreview = rep.length > 50 ? rep.slice(0, 50) + "..." : rep;
          errors.push(
            `  ${this.rel(xmlFile)}: Line ${lineOf(instrElem)}: <w:instrText> found within <w:del> (use <w:delInstrText>): ${textPreview}`
          );
        }
      } catch (e) {
        errors.push(`  ${this.rel(xmlFile)}: Error: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} deletion validation violations:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - No w:t elements found within w:del elements");
    return true;
  }

  countParagraphsInUnpacked() {
    let count = 0;
    for (const xmlFile of this.xml_files) {
      if (path.basename(xmlFile) !== "document.xml") continue;
      try {
        const root = this._parseOrThrow(xmlFile).documentElement;
        count = findDescendants(root, DOCXSchemaValidator.WORD_2006_NAMESPACE, "p").length;
      } catch (e) {
        console.log(`Error counting paragraphs in unpacked document: ${e.message}`);
      }
    }
    return count;
  }

  countParagraphsInOriginal() {
    const original = this.original_file;
    if (original === null) return 0;

    let count = 0;
    let tempDir;
    try {
      tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "docx-orig-"));
      const zip = new AdmZip(original);
      zip.extractAllTo(tempDir, true);

      const docXmlPath = tempDir + "/word/document.xml";
      const root = minidom.parseString(fs.readFileSync(docXmlPath, "utf-8")).documentElement;
      count = findDescendants(root, DOCXSchemaValidator.WORD_2006_NAMESPACE, "p").length;
    } catch (e) {
      console.log(`Error counting paragraphs in original document: ${e.message}`);
    } finally {
      if (tempDir) {
        try {
          fs.rmSync(tempDir, { recursive: true, force: true });
        } catch (e) {}
      }
    }
    return count;
  }

  validateInsertions() {
    const errors = [];
    const W = DOCXSchemaValidator.WORD_2006_NAMESPACE;

    for (const xmlFile of this.xml_files) {
      if (path.basename(xmlFile) !== "document.xml") continue;

      try {
        const root = this._parseOrThrow(xmlFile).documentElement;

        for (const elem of iterElements(root)) {
          if (!matchNs(elem, W, "delText")) continue;
          if (!hasAncestor(elem, W, "ins")) continue;
          if (hasAncestor(elem, W, "del")) continue;

          const t = elemText(elem) || "";
          const rep = pyReprStr(t);
          const textPreview = rep.length > 50 ? rep.slice(0, 50) + "..." : rep;
          errors.push(
            `  ${this.rel(xmlFile)}: Line ${lineOf(elem)}: <w:delText> within <w:ins>: ${textPreview}`
          );
        }
      } catch (e) {
        errors.push(`  ${this.rel(xmlFile)}: Error: ${e.message}`);
      }
    }

    if (errors.length) {
      console.log(`FAILED - Found ${errors.length} insertion validation violations:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - No w:delText elements within w:ins elements");
    return true;
  }

  compareParagraphCounts() {
    const originalCount = this.countParagraphsInOriginal();
    const newCount = this.countParagraphsInUnpacked();

    const diff = newCount - originalCount;
    const diffStr = diff > 0 ? `+${diff}` : String(diff);
    console.log(`\nParagraphs: ${originalCount} → ${newCount} (${diffStr})`);
  }

  _parseIdValue(val, base) {
    base = base || 16;
    const s = String(val).trim();
    let body = s;
    let sign = 1;
    if (body.startsWith("+")) body = body.slice(1);
    else if (body.startsWith("-")) {
      sign = -1;
      body = body.slice(1);
    }
    if (base === 16) {
      if (/^0[xX]/.test(body)) body = body.slice(2);
      if (!/^[0-9a-fA-F]+$/.test(body)) {
        throw new Error(`invalid literal for int() with base 16: ${pyReprStr(s)}`);
      }
      return sign * parseInt(body, 16);
    }
    if (!/^[0-9]+$/.test(body)) {
      throw new Error(`invalid literal for int() with base 10: ${pyReprStr(s)}`);
    }
    return sign * parseInt(body, 10);
  }

  validateIdConstraints() {
    const errors = [];
    const paraNs = DOCXSchemaValidator.W14_NAMESPACE;
    const durableNs = DOCXSchemaValidator.W16CID_NAMESPACE;

    for (const xmlFile of this.xml_files) {
      try {
        const root = this._parseOrThrow(xmlFile).documentElement;
        for (const elem of iterElements(root)) {
          let val = getNs(elem, paraNs, "paraId");
          if (val) {
            if (this._parseIdValue(val, 16) >= 0x80000000) {
              errors.push(`  ${path.basename(xmlFile)}:${lineOf(elem)}: paraId=${val} >= 0x80000000`);
            }
          }

          val = getNs(elem, durableNs, "durableId");
          if (val) {
            if (path.basename(xmlFile) === "numbering.xml") {
              try {
                if (this._parseIdValue(val, 10) >= 0x7fffffff) {
                  errors.push(`  ${path.basename(xmlFile)}:${lineOf(elem)}: durableId=${val} >= 0x7FFFFFFF`);
                }
              } catch (ve) {
                errors.push(`  ${path.basename(xmlFile)}:${lineOf(elem)}: durableId=${val} must be decimal in numbering.xml`);
              }
            } else {
              if (this._parseIdValue(val, 16) >= 0x7fffffff) {
                errors.push(`  ${path.basename(xmlFile)}:${lineOf(elem)}: durableId=${val} >= 0x7FFFFFFF`);
              }
            }
          }
        }
      } catch (e) {
        // pass
      }
    }

    if (errors.length) {
      console.log(`FAILED - ${errors.length} ID constraint violations:`);
      for (const e of errors) console.log(e);
    } else if (this.verbose) {
      console.log("PASSED - All paraId/durableId values within constraints");
    }
    return errors.length === 0;
  }

  validateCommentMarkers() {
    const errors = [];
    const W = DOCXSchemaValidator.WORD_2006_NAMESPACE;

    let documentXml = null;
    let commentsXml = null;
    for (const xmlFile of this.xml_files) {
      if (path.basename(xmlFile) === "document.xml" && xmlFile.split(path.sep).join("/").indexOf("word") !== -1) {
        documentXml = xmlFile;
      } else if (path.basename(xmlFile) === "comments.xml") {
        commentsXml = xmlFile;
      }
    }

    if (!documentXml) {
      if (this.verbose) console.log("PASSED - No document.xml found (skipping comment validation)");
      return true;
    }

    const sortByNumeric = (arr) =>
      arr.slice().sort((a, b) => {
        const ka = a && /^\d+$/.test(a) ? parseInt(a, 10) : 0;
        const kb = b && /^\d+$/.test(b) ? parseInt(b, 10) : 0;
        return ka - kb;
      });
    const setDiff = (a, b) => [...a].filter((x) => !b.has(x));

    try {
      const docRoot = this._parseOrThrow(documentXml).documentElement;

      const rangeStarts = new Set(
        findDescendants(docRoot, W, "commentRangeStart").map((el) => getNs(el, W, "id"))
      );
      const rangeEnds = new Set(
        findDescendants(docRoot, W, "commentRangeEnd").map((el) => getNs(el, W, "id"))
      );
      const references = new Set(
        findDescendants(docRoot, W, "commentReference").map((el) => getNs(el, W, "id"))
      );

      const orphanedEnds = new Set(setDiff(rangeEnds, rangeStarts));
      for (const commentId of sortByNumeric([...orphanedEnds])) {
        errors.push(`  document.xml: commentRangeEnd id="${commentId}" has no matching commentRangeStart`);
      }

      const orphanedStarts = new Set(setDiff(rangeStarts, rangeEnds));
      for (const commentId of sortByNumeric([...orphanedStarts])) {
        errors.push(`  document.xml: commentRangeStart id="${commentId}" has no matching commentRangeEnd`);
      }

      let commentIds = new Set();
      if (commentsXml && pyfs.exists(commentsXml)) {
        const commentsRoot = this._parseOrThrow(commentsXml).documentElement;
        commentIds = new Set(findDescendants(commentsRoot, W, "comment").map((el) => getNs(el, W, "id")));

        const markerIds = new Set([...rangeStarts, ...rangeEnds, ...references]);
        const invalidRefs = setDiff(markerIds, commentIds);
        for (const commentId of sortByNumeric(invalidRefs)) {
          if (commentId) {
            errors.push(`  document.xml: marker id="${commentId}" references non-existent comment`);
          }
        }
      }
    } catch (e) {
      errors.push(`  Error parsing XML: ${e.message}`);
    }

    if (errors.length) {
      console.log(`FAILED - ${errors.length} comment marker violations:`);
      for (const error of errors) console.log(error);
      return false;
    }
    if (this.verbose) console.log("PASSED - All comment markers properly paired");
    return true;
  }

  repair() {
    let repairs = super.repair();
    repairs += this.repairDurableId();
    return repairs;
  }

  repairDurableId() {
    let repairs = 0;

    for (const xmlFile of this.xml_files) {
      try {
        const content = fs.readFileSync(xmlFile, "utf-8");
        const dom = minidom.parseString(content);
        let modified = false;

        const all = dom.getElementsByTagName("*");
        for (let i = 0; i < all.length; i++) {
          const elem = all[i];
          if (!elem.hasAttribute("w16cid:durableId")) continue;

          const durableId = elem.getAttribute("w16cid:durableId");
          let needsRepair = false;

          if (path.basename(xmlFile) === "numbering.xml") {
            try {
              needsRepair = this._parseIdValue(durableId, 10) >= 0x7fffffff;
            } catch (ve) {
              needsRepair = true;
            }
          } else {
            try {
              needsRepair = this._parseIdValue(durableId, 16) >= 0x7fffffff;
            } catch (ve) {
              needsRepair = true;
            }
          }

          if (needsRepair) {
            const value = Math.floor(Math.random() * 0x7ffffffe) + 1;
            let newId;
            if (path.basename(xmlFile) === "numbering.xml") {
              newId = String(value);
            } else {
              newId = value.toString(16).toUpperCase().padStart(8, "0");
            }

            elem.setAttribute("w16cid:durableId", newId);
            console.log(`  Repaired: ${path.basename(xmlFile)}: durableId ${durableId} → ${newId}`);
            repairs += 1;
            modified = true;
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
}

DOCXSchemaValidator.WORD_2006_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
DOCXSchemaValidator.W14_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
DOCXSchemaValidator.W16CID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
DOCXSchemaValidator.ELEMENT_RELATIONSHIP_TYPES = {};

module.exports = { DOCXSchemaValidator };
