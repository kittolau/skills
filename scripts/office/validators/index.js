"use strict";
/*
 * Validation modules for Word document processing.
 * (Port of validators/__init__.py)
 */

const { BaseSchemaValidator } = require("./base");
const { DOCXSchemaValidator } = require("./docx");
const { PPTXSchemaValidator } = require("./pptx");
const { RedliningValidator } = require("./redlining");

module.exports = {
  BaseSchemaValidator,
  DOCXSchemaValidator,
  PPTXSchemaValidator,
  RedliningValidator,
};
