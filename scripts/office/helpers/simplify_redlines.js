"use strict";
/*
 * Simplify tracked changes by merging adjacent w:ins or w:del elements.
 * (Port of simplify_redlines.py)
 */

const fs = require("fs");
const path = require("path");
const AdmZip = require("adm-zip");
const minidom = require("../../_pylib/minidom");

const ELEMENT_NODE = minidom.ELEMENT_NODE;
const TEXT_NODE = minidom.TEXT_NODE;

const WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function localName(node) {
  return node.localName || node.tagName;
}

function simplifyRedlines(inputDir) {
  const docXml = path.join(inputDir, "word", "document.xml");

  if (!fs.existsSync(docXml)) {
    return [0, `Error: ${docXml} not found`];
  }

  try {
    const dom = minidom.parseString(fs.readFileSync(docXml, "utf-8"));
    const root = dom.documentElement;

    let mergeCount = 0;

    const containers = findElements(root, "p").concat(findElements(root, "tc"));

    for (const container of containers) {
      mergeCount += mergeTrackedChangesIn(container, "ins");
      mergeCount += mergeTrackedChangesIn(container, "del");
    }

    fs.writeFileSync(docXml, minidom.toxml(dom, "UTF-8"));
    return [mergeCount, `Simplified ${mergeCount} tracked changes`];
  } catch (e) {
    return [0, `Error: ${e.message}`];
  }
}

function mergeTrackedChangesIn(container, tag) {
  let mergeCount = 0;

  const tracked = [];
  for (const child of minidom.childNodesArray(container)) {
    if (child.nodeType === ELEMENT_NODE && isElement(child, tag)) {
      tracked.push(child);
    }
  }

  if (tracked.length < 2) {
    return 0;
  }

  let i = 0;
  while (i < tracked.length - 1) {
    const curr = tracked[i];
    const nextElem = tracked[i + 1];

    if (canMergeTracked(curr, nextElem)) {
      mergeTrackedContent(curr, nextElem);
      container.removeChild(nextElem);
      tracked.splice(i + 1, 1);
      mergeCount += 1;
    } else {
      i += 1;
    }
  }

  return mergeCount;
}

function isElement(node, tag) {
  const name = localName(node);
  return name === tag || name.endsWith(`:${tag}`);
}

function getAuthor(elem) {
  let author = elem.getAttribute("w:author");
  if (!author) {
    const attrs = elem.attributes;
    for (let i = 0; i < attrs.length; i++) {
      const attr = attrs.item(i);
      if (attr.localName === "author" || attr.name.endsWith(":author")) {
        return attr.value;
      }
    }
  }
  return author;
}

function canMergeTracked(elem1, elem2) {
  if (getAuthor(elem1) !== getAuthor(elem2)) {
    return false;
  }

  let node = elem1.nextSibling;
  while (node && node !== elem2) {
    if (node.nodeType === ELEMENT_NODE) return false;
    if (node.nodeType === TEXT_NODE && node.data.trim()) return false;
    node = node.nextSibling;
  }

  return true;
}

function mergeTrackedContent(target, source) {
  while (source.firstChild) {
    const child = source.firstChild;
    source.removeChild(child);
    target.appendChild(child);
  }
}

function findElements(root, tag) {
  const results = [];

  function traverse(node) {
    if (node.nodeType === ELEMENT_NODE) {
      const name = localName(node);
      if (name === tag || name.endsWith(`:${tag}`)) {
        results.push(node);
      }
      for (const child of minidom.childNodesArray(node)) {
        traverse(child);
      }
    }
  }

  traverse(root);
  return results;
}

// --- Author inference (uses namespace-aware traversal, mirrors ElementTree) ---

function countAuthorsInRoot(root) {
  const authors = {};
  for (const tag of ["ins", "del"]) {
    for (const elem of findNsElements(root, tag)) {
      const author = getAttrNs(elem, WORD_NS, "author");
      if (author) {
        authors[author] = (authors[author] || 0) + 1;
      }
    }
  }
  return authors;
}

function findNsElements(root, localTag) {
  const results = [];
  function traverse(node) {
    if (node.nodeType === ELEMENT_NODE) {
      if (
        (node.localName || node.tagName) === localTag &&
        node.namespaceURI === WORD_NS
      ) {
        results.push(node);
      }
      for (const child of minidom.childNodesArray(node)) traverse(child);
    }
  }
  // root itself can match in ElementTree's .// (descendants only); start at children
  for (const child of minidom.childNodesArray(root)) traverse(child);
  return results;
}

function getAttrNs(elem, ns, local) {
  if (elem.getAttributeNS) {
    const v = elem.getAttributeNS(ns, local);
    if (v) return v;
  }
  return null;
}

function getTrackedChangeAuthors(docXmlPath) {
  if (!fs.existsSync(docXmlPath)) {
    return {};
  }
  let dom;
  try {
    dom = minidom.parseString(fs.readFileSync(docXmlPath, "utf-8"));
  } catch (e) {
    return {};
  }
  return countAuthorsInRoot(dom.documentElement);
}

function getAuthorsFromDocx(docxPath) {
  try {
    const zip = new AdmZip(docxPath);
    const entry = zip.getEntry("word/document.xml");
    if (!entry) return {};
    const content = zip.readAsText(entry, "utf8");
    const dom = minidom.parseString(content);
    return countAuthorsInRoot(dom.documentElement);
  } catch (e) {
    return {};
  }
}

function inferAuthor(modifiedDir, originalDocx, defaultName) {
  if (defaultName === undefined) defaultName = "Claude";
  const modifiedXml = path.join(modifiedDir, "word", "document.xml");
  const modifiedAuthors = getTrackedChangeAuthors(modifiedXml);

  if (Object.keys(modifiedAuthors).length === 0) {
    return defaultName;
  }

  const originalAuthors = getAuthorsFromDocx(originalDocx);

  const newChanges = {};
  for (const author of Object.keys(modifiedAuthors)) {
    const count = modifiedAuthors[author];
    const originalCount = originalAuthors[author] || 0;
    const diff = count - originalCount;
    if (diff > 0) {
      newChanges[author] = diff;
    }
  }

  const keys = Object.keys(newChanges);
  if (keys.length === 0) {
    return defaultName;
  }

  if (keys.length === 1) {
    return keys[0];
  }

  throw new Error(
    `Multiple authors added new changes: ${JSON.stringify(newChanges)}. ` +
      "Cannot infer which author to validate."
  );
}

module.exports = {
  simplifyRedlines,
  getTrackedChangeAuthors,
  inferAuthor,
};
