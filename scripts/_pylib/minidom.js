"use strict";
/*
 * minidom.js
 *
 * A faithful re-implementation of the parts of Python's `xml.dom.minidom`
 * (and `defusedxml.minidom`) that the ported scripts rely on.
 *
 * The goal is byte-for-byte identical serialization output compared to
 * CPython's minidom, because the original Python scripts rewrite XML files
 * in place with `dom.toxml(encoding=...)` / `dom.toprettyxml(indent=..., encoding=...)`.
 *
 * Parsing is delegated to @xmldom/xmldom (pure JS, DOM Level 2 compatible),
 * and serialization is implemented here to mirror CPython's writexml() logic.
 */

const fs = require("fs");
const { DOMParser } = require("@xmldom/xmldom");

// DOM node type constants (same numeric values as the W3C DOM / Python minidom)
const ELEMENT_NODE = 1;
const ATTRIBUTE_NODE = 2;
const TEXT_NODE = 3;
const CDATA_SECTION_NODE = 4;
const PROCESSING_INSTRUCTION_NODE = 7;
const COMMENT_NODE = 8;
const DOCUMENT_NODE = 9;
const DOCUMENT_TYPE_NODE = 10;

// Mirror of CPython minidom._write_data: escapes &, <, ", > (in that order).
function writeData(data) {
  if (!data) return "";
  return data
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/"/g, "&quot;")
    .replace(/>/g, "&gt;");
}

// Strict-ish error handler so malformed XML raises (mirrors expat raising).
function makeParser() {
  return new DOMParser({
    locator: {},
    errorHandler: {
      warning: function () {},
      error: function (msg) {
        throw new Error(msg);
      },
      fatalError: function (msg) {
        throw new Error(msg);
      },
    },
  });
}

function parseString(text) {
  const parser = makeParser();
  const doc = parser.parseFromString(text, "text/xml");
  return doc;
}

function parseFile(path) {
  const text = fs.readFileSync(path, "utf-8");
  return parseString(text);
}

// ---------------------------------------------------------------------------
// Serialization (mirrors CPython xml.dom.minidom writexml methods)
// ---------------------------------------------------------------------------

function elementAttrNames(elem) {
  // Returns attribute qualified names in document (parse) order, matching
  // CPython 3.8+ minidom which preserves attribute order.
  const names = [];
  const attrs = elem.attributes;
  if (!attrs) return names;
  for (let i = 0; i < attrs.length; i++) {
    names.push(attrs.item(i).name);
  }
  return names;
}

function getAttrValue(elem, name) {
  const attrs = elem.attributes;
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item(i);
    if (a.name === name) return a.value;
  }
  return "";
}

function childNodesArray(node) {
  const out = [];
  const cn = node.childNodes;
  if (!cn) return out;
  for (let i = 0; i < cn.length; i++) out.push(cn.item ? cn.item(i) : cn[i]);
  return out;
}

// True for the XML declaration, which @xmldom exposes as a processing
// instruction with target "xml" but CPython minidom never keeps as a child.
function isXmlDeclaration(node) {
  return (
    node.nodeType === PROCESSING_INSTRUCTION_NODE &&
    String(node.target).toLowerCase() === "xml"
  );
}

function writeNode(node, out, indent, addindent, newl) {
  switch (node.nodeType) {
    case ELEMENT_NODE:
      return writeElement(node, out, indent, addindent, newl);
    case TEXT_NODE:
      // Text.writexml: _write_data(writer, indent + data + newl)
      out.push(writeData(indent + (node.data || "") + newl));
      return;
    case CDATA_SECTION_NODE:
      out.push("<![CDATA[" + (node.data || "") + "]]>");
      return;
    case COMMENT_NODE:
      out.push(indent + "<!--" + (node.data || "") + "-->" + newl);
      return;
    case PROCESSING_INSTRUCTION_NODE:
      out.push(indent + "<?" + node.target + " " + (node.data || "") + "?>" + newl);
      return;
    case DOCUMENT_TYPE_NODE:
      writeDocumentType(node, out, newl);
      return;
    default:
      // Unknown node types are ignored, matching practical minidom usage here.
      return;
  }
}

function writeElement(elem, out, indent, addindent, newl) {
  out.push(indent + "<" + elem.tagName);

  const names = elementAttrNames(elem);
  for (const aName of names) {
    out.push(' ' + aName + '="');
    out.push(writeData(getAttrValue(elem, aName)));
    out.push('"');
  }

  const children = childNodesArray(elem);
  if (children.length > 0) {
    out.push(">");
    if (
      children.length === 1 &&
      (children[0].nodeType === TEXT_NODE ||
        children[0].nodeType === CDATA_SECTION_NODE)
    ) {
      // Single text/CDATA child is written inline with no indent/newl.
      writeNode(children[0], out, "", "", "");
    } else {
      out.push(newl);
      for (const child of children) {
        writeNode(child, out, indent + addindent, addindent, newl);
      }
      out.push(indent);
    }
    out.push("</" + elem.tagName + ">" + newl);
  } else {
    out.push("/>" + newl);
  }
}

function writeDocumentType(node, out, newl) {
  // Minimal DOCTYPE serialization; office documents generally have none.
  const name = node.name || node.nodeName || "";
  out.push("<!DOCTYPE " + name);
  if (node.publicId) {
    out.push(' PUBLIC "' + node.publicId + '"');
    out.push(' "' + (node.systemId || "") + '"');
  } else if (node.systemId) {
    out.push(' SYSTEM "' + node.systemId + '"');
  }
  out.push(">" + newl);
}

function xmlDeclaration(encoding, standalone, newl) {
  const encStr = encoding == null ? "" : ' encoding="' + encoding + '"';
  let saStr = "";
  if (standalone != null) {
    saStr = ' standalone="' + (standalone ? "yes" : "no") + '"';
  }
  return '<?xml version="1.0"' + encStr + saStr + "?>" + newl;
}

/*
 * toprettyxml(node, {indent, newl, encoding, standalone})
 * Mirrors CPython Node.toprettyxml. For a Document node it emits the XML
 * declaration; for any other node it serializes that node only.
 * Returns a string (callers write it as UTF-8 bytes).
 */
function toprettyxml(node, options) {
  options = options || {};
  const indent = options.indent != null ? options.indent : "\t";
  const newl = options.newl != null ? options.newl : "\n";
  const encoding = options.encoding != null ? options.encoding : null;
  const standalone = options.standalone != null ? options.standalone : null;

  const out = [];
  if (node.nodeType === DOCUMENT_NODE) {
    out.push(xmlDeclaration(encoding, standalone, newl));
    for (const child of childNodesArray(node)) {
      if (isXmlDeclaration(child)) continue; // never re-emit the source decl
      writeNode(child, out, "", indent, newl);
    }
  } else {
    writeNode(node, out, "", indent, newl);
  }
  return out.join("");
}

/*
 * toxml(node, encoding) -> string
 * Equivalent to CPython Node.toxml(encoding): toprettyxml("", "", encoding).
 */
function toxml(node, encoding) {
  return toprettyxml(node, {
    indent: "",
    newl: "",
    encoding: encoding != null ? encoding : null,
  });
}

module.exports = {
  ELEMENT_NODE,
  ATTRIBUTE_NODE,
  TEXT_NODE,
  CDATA_SECTION_NODE,
  PROCESSING_INSTRUCTION_NODE,
  COMMENT_NODE,
  DOCUMENT_NODE,
  DOCUMENT_TYPE_NODE,
  parseString,
  parseFile,
  toxml,
  toprettyxml,
  writeData,
  childNodesArray,
};
