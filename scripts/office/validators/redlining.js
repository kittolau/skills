"use strict";
/*
 * Validator for tracked changes in Word documents.
 * (Port of validators/redlining.py)
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const AdmZip = require("adm-zip");
const minidom = require("../../_pylib/minidom");
const pyfs = require("../../_pylib/pyfs");

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function localOf(node) {
  return node.localName || node.tagName;
}
function matchNs(el, ns, local) {
  return el.nodeType === minidom.ELEMENT_NODE && localOf(el) === local && el.namespaceURI === ns;
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
function childElements(parent) {
  return minidom.childNodesArray(parent).filter((n) => n.nodeType === minidom.ELEMENT_NODE);
}
function findDescendants(root, ns, local) {
  return iterElements(root).filter((el) => el !== root && matchNs(el, ns, local));
}
function getNs(elem, ns, local) {
  if (elem.getAttributeNS) {
    const v = elem.getAttributeNS(ns, local);
    if (v !== "" || (elem.hasAttributeNS && elem.hasAttributeNS(ns, local))) return v;
  }
  return null;
}
function elemText(el) {
  const first = el.firstChild;
  if (first && first.nodeType === minidom.TEXT_NODE) return first.data;
  return null;
}

class RedliningValidator {
  constructor(unpackedDir, originalDocx, verbose, author) {
    this.unpacked_dir = unpackedDir;
    this.original_docx = originalDocx;
    this.verbose = !!verbose;
    this.author = author != null ? author : "Claude";
  }

  repair() {
    return 0;
  }

  validate() {
    const modifiedFile = path.join(this.unpacked_dir, "word", "document.xml");
    if (!pyfs.exists(modifiedFile)) {
      console.log(`FAILED - Modified document.xml not found at ${modifiedFile}`);
      return false;
    }

    try {
      const root = minidom.parseString(fs.readFileSync(modifiedFile, "utf-8")).documentElement;

      const delElements = findDescendants(root, W, "del");
      const insElements = findDescendants(root, W, "ins");

      const authorDel = delElements.filter((elem) => getNs(elem, W, "author") === this.author);
      const authorIns = insElements.filter((elem) => getNs(elem, W, "author") === this.author);

      if (authorDel.length === 0 && authorIns.length === 0) {
        if (this.verbose) console.log(`PASSED - No tracked changes by ${this.author} found.`);
        return true;
      }
    } catch (e) {
      // pass
    }

    const tempPath = fs.mkdtempSync(path.join(os.tmpdir(), "redline-"));
    try {
      try {
        const zip = new AdmZip(this.original_docx);
        zip.extractAllTo(tempPath, true);
      } catch (e) {
        console.log(`FAILED - Error unpacking original docx: ${e.message}`);
        return false;
      }

      const originalFile = path.join(tempPath, "word", "document.xml");
      if (!pyfs.exists(originalFile)) {
        console.log(`FAILED - Original document.xml not found in ${this.original_docx}`);
        return false;
      }

      let modifiedRoot, originalRoot;
      try {
        const modifiedFileNode = path.join(this.unpacked_dir, "word", "document.xml");
        modifiedRoot = minidom.parseString(fs.readFileSync(modifiedFileNode, "utf-8")).documentElement;
        originalRoot = minidom.parseString(fs.readFileSync(originalFile, "utf-8")).documentElement;
      } catch (e) {
        console.log(`FAILED - Error parsing XML files: ${e.message}`);
        return false;
      }

      this._removeAuthorTrackedChanges(originalRoot);
      this._removeAuthorTrackedChanges(modifiedRoot);

      const modifiedText = this._extractTextContent(modifiedRoot);
      const originalText = this._extractTextContent(originalRoot);

      if (modifiedText !== originalText) {
        const errorMessage = this._generateDetailedDiff(originalText, modifiedText);
        console.log(errorMessage);
        return false;
      }

      if (this.verbose) console.log(`PASSED - All changes by ${this.author} are properly tracked`);
      return true;
    } finally {
      try {
        fs.rmSync(tempPath, { recursive: true, force: true });
      } catch (e) {}
    }
  }

  _generateDetailedDiff(originalText, modifiedText) {
    const errorParts = [
      `FAILED - Document text doesn't match after removing ${this.author}'s tracked changes`,
      "",
      "Likely causes:",
      "  1. Modified text inside another author's <w:ins> or <w:del> tags",
      "  2. Made edits without proper tracked changes",
      "  3. Didn't nest <w:del> inside <w:ins> when deleting another's insertion",
      "",
      "For pre-redlined documents, use correct patterns:",
      "  - To reject another's INSERTION: Nest <w:del> inside their <w:ins>",
      "  - To restore another's DELETION: Add new <w:ins> AFTER their <w:del>",
      "",
    ];

    const gitDiff = this._getGitWordDiff(originalText, modifiedText);
    if (gitDiff) {
      errorParts.push("Differences:", "============", gitDiff);
    } else {
      errorParts.push("Unable to generate word diff (git not available)");
    }

    return errorParts.join("\n");
  }

  _getGitWordDiff(originalText, modifiedText) {
    try {
      const tempPath = fs.mkdtempSync(path.join(os.tmpdir(), "redline-diff-"));
      try {
        const originalFile = path.join(tempPath, "original.txt");
        const modifiedFile = path.join(tempPath, "modified.txt");

        fs.writeFileSync(originalFile, originalText, { encoding: "utf-8" });
        fs.writeFileSync(modifiedFile, modifiedText, { encoding: "utf-8" });

        const parseOut = (stdout) => {
          const lines = stdout.split("\n");
          const contentLines = [];
          let inContent = false;
          for (const line of lines) {
            if (line.startsWith("@@")) {
              inContent = true;
              continue;
            }
            if (inContent && line.trim()) contentLines.push(line);
          }
          return contentLines;
        };

        let result = spawnSync(
          "git",
          [
            "diff",
            "--word-diff=plain",
            "--word-diff-regex=.",
            "-U0",
            "--no-index",
            originalFile,
            modifiedFile,
          ],
          { encoding: "utf-8" }
        );

        if (result.stdout && result.stdout.trim()) {
          const contentLines = parseOut(result.stdout);
          if (contentLines.length) return contentLines.join("\n");
        }

        result = spawnSync(
          "git",
          ["diff", "--word-diff=plain", "-U0", "--no-index", originalFile, modifiedFile],
          { encoding: "utf-8" }
        );

        if (result.stdout && result.stdout.trim()) {
          const contentLines = parseOut(result.stdout);
          return contentLines.join("\n");
        }
      } finally {
        try {
          fs.rmSync(tempPath, { recursive: true, force: true });
        } catch (e) {}
      }
    } catch (e) {
      // pass
    }

    return null;
  }

  _removeAuthorTrackedChanges(root) {
    // Phase 1: drop author's <w:ins> elements entirely.
    for (const parent of iterElements(root)) {
      const toRemove = [];
      for (const child of childElements(parent)) {
        if (matchNs(child, W, "ins") && getNs(child, W, "author") === this.author) {
          toRemove.push(child);
        }
      }
      for (const elem of toRemove) parent.removeChild(elem);
    }

    // Phase 2: unwrap author's <w:del>, converting <w:delText> -> <w:t>.
    const authorDels = [];
    for (const parent of iterElements(root)) {
      for (const child of childElements(parent)) {
        if (matchNs(child, W, "del") && getNs(child, W, "author") === this.author) {
          authorDels.push(child);
        }
      }
    }

    for (const delElem of authorDels) {
      const parent = delElem.parentNode;
      if (!parent) continue;

      // delText -> t (replace node, preserving text child)
      for (const el of iterElements(delElem)) {
        if (matchNs(el, W, "delText")) {
          const doc = el.ownerDocument;
          const t = doc.createElementNS(W, "w:t");
          const attrs = el.attributes;
          for (let i = 0; i < attrs.length; i++) {
            const a = attrs.item(i);
            t.setAttributeNS(a.namespaceURI || null, a.name, a.value);
          }
          while (el.firstChild) t.appendChild(el.firstChild);
          el.parentNode.replaceChild(t, el);
        }
      }

      // unwrap: move del's element children before del, then remove del.
      for (const child of childElements(delElem)) {
        parent.insertBefore(child, delElem);
      }
      parent.removeChild(delElem);
    }
  }

  _extractTextContent(root) {
    const paragraphs = [];
    for (const pElem of findDescendants(root, W, "p")) {
      const textParts = [];
      for (const tElem of findDescendants(pElem, W, "t")) {
        const txt = elemText(tElem);
        if (txt) textParts.push(txt);
      }
      const paragraphText = textParts.join("");
      if (paragraphText) paragraphs.push(paragraphText);
    }
    return paragraphs.join("\n");
  }
}

module.exports = { RedliningValidator };
