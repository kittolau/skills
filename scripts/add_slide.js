"use strict";
/*
 * Add a new slide to an unpacked PPTX directory.  (Node.js port of add_slide.py)
 *
 * Usage: node add_slide.js <unpacked_dir> <source>
 *
 * The source can be:
 *   - A slide file (e.g., slide2.xml) - duplicates the slide
 *   - A layout file (e.g., slideLayout2.xml) - creates from layout
 *
 * Prints the <p:sldId> element to add to presentation.xml.
 */

const fs = require("fs");
const path = require("path");
const pyfs = require("./_pylib/pyfs");

function getNextSlideNumber(slidesDir) {
  const existing = [];
  for (const f of pyfs.glob(slidesDir, "slide*.xml")) {
    const m = /^slide(\d+)\.xml$/.exec(path.basename(f));
    if (m) existing.push(parseInt(m[1], 10));
  }
  return existing.length ? Math.max.apply(null, existing) + 1 : 1;
}

function createSlideFromLayout(unpackedDir, layoutFile) {
  const slidesDir = path.join(unpackedDir, "ppt", "slides");
  const relsDir = path.join(slidesDir, "_rels");
  const layoutsDir = path.join(unpackedDir, "ppt", "slideLayouts");

  const layoutPath = path.join(layoutsDir, layoutFile);
  if (!pyfs.exists(layoutPath)) {
    process.stderr.write(`Error: ${layoutPath} not found\n`);
    process.exit(1);
  }

  const nextNum = getNextSlideNumber(slidesDir);
  const dest = `slide${nextNum}.xml`;
  const destSlide = path.join(slidesDir, dest);
  const destRels = path.join(relsDir, `${dest}.rels`);

  const slideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`;
  fs.writeFileSync(destSlide, slideXml, { encoding: "utf-8" });

  fs.mkdirSync(relsDir, { recursive: true });
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/${layoutFile}"/>
</Relationships>`;
  fs.writeFileSync(destRels, relsXml, { encoding: "utf-8" });

  addToContentTypes(unpackedDir, dest);

  const rid = addToPresentationRels(unpackedDir, dest);

  const nextSlideId = getNextSlideId(unpackedDir);

  console.log(`Created ${dest} from ${layoutFile}`);
  console.log(
    `Add to presentation.xml <p:sldIdLst>: <p:sldId id="${nextSlideId}" r:id="${rid}"/>`
  );
}

function duplicateSlide(unpackedDir, source) {
  const slidesDir = path.join(unpackedDir, "ppt", "slides");
  const relsDir = path.join(slidesDir, "_rels");

  const sourceSlide = path.join(slidesDir, source);

  if (!pyfs.exists(sourceSlide)) {
    process.stderr.write(`Error: ${sourceSlide} not found\n`);
    process.exit(1);
  }

  const nextNum = getNextSlideNumber(slidesDir);
  const dest = `slide${nextNum}.xml`;
  const destSlide = path.join(slidesDir, dest);

  const sourceRels = path.join(relsDir, `${source}.rels`);
  const destRels = path.join(relsDir, `${dest}.rels`);

  fs.copyFileSync(sourceSlide, destSlide);

  if (pyfs.exists(sourceRels)) {
    fs.copyFileSync(sourceRels, destRels);

    let relsContent = fs.readFileSync(destRels, "utf-8");
    relsContent = relsContent.replace(
      /\s*<Relationship[^>]*Type="[^"]*notesSlide"[^>]*\/>\s*/g,
      "\n"
    );
    fs.writeFileSync(destRels, relsContent, { encoding: "utf-8" });
  }

  addToContentTypes(unpackedDir, dest);

  const rid = addToPresentationRels(unpackedDir, dest);

  const nextSlideId = getNextSlideId(unpackedDir);

  console.log(`Created ${dest} from ${source}`);
  console.log(
    `Add to presentation.xml <p:sldIdLst>: <p:sldId id="${nextSlideId}" r:id="${rid}"/>`
  );
}

function addToContentTypes(unpackedDir, dest) {
  const contentTypesPath = path.join(unpackedDir, "[Content_Types].xml");
  let contentTypes = fs.readFileSync(contentTypesPath, "utf-8");

  const newOverride = `<Override PartName="/ppt/slides/${dest}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`;

  if (contentTypes.indexOf(`/ppt/slides/${dest}`) === -1) {
    contentTypes = contentTypes.replace(
      "</Types>",
      `  ${newOverride}\n</Types>`
    );
    fs.writeFileSync(contentTypesPath, contentTypes, { encoding: "utf-8" });
  }
}

function addToPresentationRels(unpackedDir, dest) {
  const presRelsPath = path.join(
    unpackedDir,
    "ppt",
    "_rels",
    "presentation.xml.rels"
  );
  let presRels = fs.readFileSync(presRelsPath, "utf-8");

  const rids = [];
  let m;
  const re = /Id="rId(\d+)"/g;
  while ((m = re.exec(presRels)) !== null) rids.push(parseInt(m[1], 10));
  const nextRid = rids.length ? Math.max.apply(null, rids) + 1 : 1;
  const rid = `rId${nextRid}`;

  const newRel = `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/${dest}"/>`;

  if (presRels.indexOf(`slides/${dest}`) === -1) {
    presRels = presRels.replace(
      "</Relationships>",
      `  ${newRel}\n</Relationships>`
    );
    fs.writeFileSync(presRelsPath, presRels, { encoding: "utf-8" });
  }

  return rid;
}

function getNextSlideId(unpackedDir) {
  const presPath = path.join(unpackedDir, "ppt", "presentation.xml");
  const presContent = fs.readFileSync(presPath, "utf-8");
  const slideIds = [];
  let m;
  const re = /<p:sldId[^>]*id="(\d+)"/g;
  while ((m = re.exec(presContent)) !== null) slideIds.push(parseInt(m[1], 10));
  return slideIds.length ? Math.max.apply(null, slideIds) + 1 : 256;
}

function parseSource(source) {
  if (source.startsWith("slideLayout") && source.endsWith(".xml")) {
    return ["layout", source];
  }
  return ["slide", null];
}

function main() {
  const argv = process.argv.slice(2);
  if (argv.length !== 2) {
    process.stderr.write("Usage: node add_slide.js <unpacked_dir> <source>\n");
    process.stderr.write("\n");
    process.stderr.write("Source can be:\n");
    process.stderr.write("  slide2.xml        - duplicate an existing slide\n");
    process.stderr.write("  slideLayout2.xml  - create from a layout template\n");
    process.stderr.write("\n");
    process.stderr.write(
      "To see available layouts: ls <unpacked_dir>/ppt/slideLayouts/\n"
    );
    process.exit(1);
  }

  const unpackedDir = argv[0];
  const source = argv[1];

  if (!pyfs.exists(unpackedDir)) {
    process.stderr.write(`Error: ${unpackedDir} not found\n`);
    process.exit(1);
  }

  const [sourceType, layoutFile] = parseSource(source);

  if (sourceType === "layout" && layoutFile !== null) {
    createSlideFromLayout(unpackedDir, layoutFile);
  } else {
    duplicateSlide(unpackedDir, source);
  }
}

if (require.main === module) {
  main();
}

module.exports = {
  getNextSlideNumber,
  createSlideFromLayout,
  duplicateSlide,
  parseSource,
};
