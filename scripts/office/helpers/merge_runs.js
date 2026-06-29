"use strict";
/*
 * Merge adjacent runs with identical formatting in DOCX. (Port of merge_runs.py)
 */

const fs = require("fs");
const path = require("path");
const minidom = require("../../_pylib/minidom");

const ELEMENT_NODE = minidom.ELEMENT_NODE;
const TEXT_NODE = minidom.TEXT_NODE;

function localName(node) {
  return node.localName || node.tagName;
}

function mergeRuns(inputDir) {
  const docXml = path.join(inputDir, "word", "document.xml");

  if (!fs.existsSync(docXml)) {
    return [0, `Error: ${docXml} not found`];
  }

  try {
    const dom = minidom.parseString(fs.readFileSync(docXml, "utf-8"));
    const root = dom.documentElement;

    removeElements(root, "proofErr");
    stripRunRsidAttrs(root);

    const containers = new Set();
    for (const run of findElements(root, "r")) {
      containers.add(run.parentNode);
    }

    let mergeCount = 0;
    for (const container of containers) {
      mergeCount += mergeRunsIn(container);
    }

    fs.writeFileSync(docXml, minidom.toxml(dom, "UTF-8"));
    return [mergeCount, `Merged ${mergeCount} runs`];
  } catch (e) {
    return [0, `Error: ${e.message}`];
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

function getChild(parent, tag) {
  for (const child of minidom.childNodesArray(parent)) {
    if (child.nodeType === ELEMENT_NODE) {
      const name = localName(child);
      if (name === tag || name.endsWith(`:${tag}`)) {
        return child;
      }
    }
  }
  return null;
}

function getChildren(parent, tag) {
  const results = [];
  for (const child of minidom.childNodesArray(parent)) {
    if (child.nodeType === ELEMENT_NODE) {
      const name = localName(child);
      if (name === tag || name.endsWith(`:${tag}`)) {
        results.push(child);
      }
    }
  }
  return results;
}

function isAdjacent(elem1, elem2) {
  let node = elem1.nextSibling;
  while (node) {
    if (node === elem2) return true;
    if (node.nodeType === ELEMENT_NODE) return false;
    if (node.nodeType === TEXT_NODE && node.data.trim()) return false;
    node = node.nextSibling;
  }
  return false;
}

function removeElements(root, tag) {
  for (const elem of findElements(root, tag)) {
    if (elem.parentNode) {
      elem.parentNode.removeChild(elem);
    }
  }
}

function stripRunRsidAttrs(root) {
  for (const run of findElements(root, "r")) {
    const attrs = run.attributes;
    const snapshot = [];
    for (let i = 0; i < attrs.length; i++) snapshot.push(attrs.item(i));
    for (const attr of snapshot) {
      if (attr.name.toLowerCase().indexOf("rsid") !== -1) {
        run.removeAttribute(attr.name);
      }
    }
  }
}

function mergeRunsIn(container) {
  let mergeCount = 0;
  let run = firstChildRun(container);

  while (run) {
    while (true) {
      const nextElem = nextElementSibling(run);
      if (nextElem && isRun(nextElem) && canMerge(run, nextElem)) {
        mergeRunContent(run, nextElem);
        container.removeChild(nextElem);
        mergeCount += 1;
      } else {
        break;
      }
    }

    consolidateText(run);
    run = nextSiblingRun(run);
  }

  return mergeCount;
}

function firstChildRun(container) {
  for (const child of minidom.childNodesArray(container)) {
    if (child.nodeType === ELEMENT_NODE && isRun(child)) {
      return child;
    }
  }
  return null;
}

function nextElementSibling(node) {
  let sibling = node.nextSibling;
  while (sibling) {
    if (sibling.nodeType === ELEMENT_NODE) return sibling;
    sibling = sibling.nextSibling;
  }
  return null;
}

function nextSiblingRun(node) {
  let sibling = node.nextSibling;
  while (sibling) {
    if (sibling.nodeType === ELEMENT_NODE) {
      if (isRun(sibling)) return sibling;
    }
    sibling = sibling.nextSibling;
  }
  return null;
}

function isRun(node) {
  const name = localName(node);
  return name === "r" || name.endsWith(":r");
}

function canMerge(run1, run2) {
  const rpr1 = getChild(run1, "rPr");
  const rpr2 = getChild(run2, "rPr");

  if ((rpr1 === null) !== (rpr2 === null)) {
    return false;
  }
  if (rpr1 === null) {
    return true;
  }
  return minidom.toxml(rpr1) === minidom.toxml(rpr2);
}

function mergeRunContent(target, source) {
  for (const child of minidom.childNodesArray(source)) {
    if (child.nodeType === ELEMENT_NODE) {
      const name = localName(child);
      if (name !== "rPr" && !name.endsWith(":rPr")) {
        target.appendChild(child);
      }
    }
  }
}

function consolidateText(run) {
  const tElements = getChildren(run, "t");

  for (let i = tElements.length - 1; i > 0; i--) {
    const curr = tElements[i];
    const prev = tElements[i - 1];

    if (isAdjacent(prev, curr)) {
      const prevText = prev.firstChild ? prev.firstChild.data : "";
      const currText = curr.firstChild ? curr.firstChild.data : "";
      const merged = prevText + currText;

      if (prev.firstChild) {
        prev.firstChild.data = merged;
      } else {
        prev.appendChild(run.ownerDocument.createTextNode(merged));
      }

      if (merged.startsWith(" ") || merged.endsWith(" ")) {
        prev.setAttribute("xml:space", "preserve");
      } else if (prev.hasAttribute("xml:space")) {
        prev.removeAttribute("xml:space");
      }

      run.removeChild(curr);
    }
  }
}

module.exports = { mergeRuns };
