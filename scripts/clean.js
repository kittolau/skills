"use strict";
/*
 * Remove unreferenced files from an unpacked PPTX directory. (Port of clean.py)
 *
 * Usage: node clean.js <unpacked_dir>
 */

const fs = require("fs");
const path = require("path");
const pyfs = require("./_pylib/pyfs");
const minidom = require("./_pylib/minidom");

function getElementsByTagNameArray(node, tag) {
  const live = node.getElementsByTagName(tag);
  const arr = [];
  for (let i = 0; i < live.length; i++) arr.push(live[i]);
  return arr;
}

function getSlidesInSldidlst(unpackedDir) {
  const presPath = path.join(unpackedDir, "ppt", "presentation.xml");
  const presRelsPath = path.join(
    unpackedDir,
    "ppt",
    "_rels",
    "presentation.xml.rels"
  );

  if (!pyfs.exists(presPath) || !pyfs.exists(presRelsPath)) {
    return new Set();
  }

  const relsDom = minidom.parseFile(presRelsPath);
  const ridToSlide = {};
  for (const rel of getElementsByTagNameArray(relsDom, "Relationship")) {
    const rid = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    const relType = rel.getAttribute("Type");
    if (relType.indexOf("slide") !== -1 && target.startsWith("slides/")) {
      ridToSlide[rid] = target.replace("slides/", "");
    }
  }

  const presContent = fs.readFileSync(presPath, "utf-8");
  const referencedRids = new Set();
  let m;
  const re = /<p:sldId[^>]*r:id="([^"]+)"/g;
  while ((m = re.exec(presContent)) !== null) referencedRids.add(m[1]);

  const result = new Set();
  for (const rid of referencedRids) {
    if (Object.prototype.hasOwnProperty.call(ridToSlide, rid)) {
      result.add(ridToSlide[rid]);
    }
  }
  return result;
}

function removeOrphanedSlides(unpackedDir) {
  const slidesDir = path.join(unpackedDir, "ppt", "slides");
  const slidesRelsDir = path.join(slidesDir, "_rels");
  const presRelsPath = path.join(
    unpackedDir,
    "ppt",
    "_rels",
    "presentation.xml.rels"
  );

  if (!pyfs.exists(slidesDir)) {
    return [];
  }

  const referencedSlides = getSlidesInSldidlst(unpackedDir);
  const removed = [];

  for (const slideFile of pyfs.glob(slidesDir, "slide*.xml")) {
    const name = path.basename(slideFile);
    if (!referencedSlides.has(name)) {
      const relPath = pyfs.relativeTo(unpackedDir, slideFile);
      fs.unlinkSync(slideFile);
      removed.push(relPath);

      const relsFile = path.join(slidesRelsDir, `${name}.rels`);
      if (pyfs.exists(relsFile)) {
        fs.unlinkSync(relsFile);
        removed.push(pyfs.relativeTo(unpackedDir, relsFile));
      }
    }
  }

  if (removed.length && pyfs.exists(presRelsPath)) {
    const relsDom = minidom.parseFile(presRelsPath);
    let changed = false;

    for (const rel of getElementsByTagNameArray(relsDom, "Relationship")) {
      const target = rel.getAttribute("Target");
      if (target.startsWith("slides/")) {
        const slideName = target.replace("slides/", "");
        if (!referencedSlides.has(slideName)) {
          if (rel.parentNode) {
            rel.parentNode.removeChild(rel);
            changed = true;
          }
        }
      }
    }

    if (changed) {
      fs.writeFileSync(presRelsPath, minidom.toxml(relsDom, "utf-8"));
    }
  }

  return removed;
}

function removeTrashDirectory(unpackedDir) {
  const trashDir = path.join(unpackedDir, "[trash]");
  const removed = [];

  if (pyfs.exists(trashDir) && pyfs.isDir(trashDir)) {
    for (const filePath of pyfs.iterdir(trashDir)) {
      if (pyfs.isFile(filePath)) {
        const relPath = pyfs.relativeTo(unpackedDir, filePath);
        removed.push(relPath);
        fs.unlinkSync(filePath);
      }
    }
    fs.rmdirSync(trashDir);
  }

  return removed;
}

function addReferencedTargets(unpackedDir, relsFile, referenced) {
  const dom = minidom.parseFile(relsFile);
  for (const rel of getElementsByTagNameArray(dom, "Relationship")) {
    const target = rel.getAttribute("Target");
    if (!target) continue;
    // (rels_file.parent.parent / target).resolve()
    const base = path.dirname(path.dirname(relsFile));
    const targetPath = path.resolve(base, target);
    const rel2 = path.relative(path.resolve(unpackedDir), targetPath);
    if (rel2.startsWith("..") || path.isAbsolute(rel2)) {
      continue; // mirrors `except ValueError: pass`
    }
    referenced.add(rel2);
  }
}

function getSlideReferencedFiles(unpackedDir) {
  const referenced = new Set();
  const slidesRelsDir = path.join(unpackedDir, "ppt", "slides", "_rels");

  if (!pyfs.exists(slidesRelsDir)) {
    return referenced;
  }

  for (const relsFile of pyfs.glob(slidesRelsDir, "*.rels")) {
    addReferencedTargets(unpackedDir, relsFile, referenced);
  }

  return referenced;
}

function removeOrphanedRelsFiles(unpackedDir) {
  const resourceDirs = ["charts", "diagrams", "drawings"];
  const removed = [];
  const slideReferenced = getSlideReferencedFiles(unpackedDir);

  for (const dirName of resourceDirs) {
    const relsDir = path.join(unpackedDir, "ppt", dirName, "_rels");
    if (!pyfs.exists(relsDir)) {
      continue;
    }

    for (const relsFile of pyfs.glob(relsDir, "*.rels")) {
      const resourceFile = path.join(
        path.dirname(relsDir),
        path.basename(relsFile).replace(".rels", "")
      );
      const resourceRelPath = path.relative(
        path.resolve(unpackedDir),
        path.resolve(resourceFile)
      );
      if (resourceRelPath.startsWith("..") || path.isAbsolute(resourceRelPath)) {
        continue; // except ValueError: continue
      }

      if (!pyfs.exists(resourceFile) || !slideReferenced.has(resourceRelPath)) {
        fs.unlinkSync(relsFile);
        const relPath = pyfs.relativeTo(unpackedDir, relsFile);
        removed.push(relPath);
      }
    }
  }

  return removed;
}

function getReferencedFiles(unpackedDir) {
  const referenced = new Set();

  for (const relsFile of pyfs.rglob(unpackedDir, "*.rels")) {
    addReferencedTargets(unpackedDir, relsFile, referenced);
  }

  return referenced;
}

function removeOrphanedFiles(unpackedDir, referenced) {
  const resourceDirs = [
    "media",
    "embeddings",
    "charts",
    "diagrams",
    "tags",
    "drawings",
    "ink",
  ];
  const removed = [];

  for (const dirName of resourceDirs) {
    const dirPath = path.join(unpackedDir, "ppt", dirName);
    if (!pyfs.exists(dirPath)) {
      continue;
    }

    for (const filePath of pyfs.glob(dirPath, "*")) {
      if (!pyfs.isFile(filePath)) {
        continue;
      }
      const relPath = pyfs.relativeTo(unpackedDir, filePath);
      if (!referenced.has(relPath)) {
        fs.unlinkSync(filePath);
        removed.push(relPath);
      }
    }
  }

  const themeDir = path.join(unpackedDir, "ppt", "theme");
  if (pyfs.exists(themeDir)) {
    for (const filePath of pyfs.glob(themeDir, "theme*.xml")) {
      const relPath = pyfs.relativeTo(unpackedDir, filePath);
      if (!referenced.has(relPath)) {
        fs.unlinkSync(filePath);
        removed.push(relPath);
        const themeRels = path.join(
          themeDir,
          "_rels",
          `${path.basename(filePath)}.rels`
        );
        if (pyfs.exists(themeRels)) {
          fs.unlinkSync(themeRels);
          removed.push(pyfs.relativeTo(unpackedDir, themeRels));
        }
      }
    }
  }

  const notesDir = path.join(unpackedDir, "ppt", "notesSlides");
  if (pyfs.exists(notesDir)) {
    for (const filePath of pyfs.glob(notesDir, "*.xml")) {
      if (!pyfs.isFile(filePath)) {
        continue;
      }
      const relPath = pyfs.relativeTo(unpackedDir, filePath);
      if (!referenced.has(relPath)) {
        fs.unlinkSync(filePath);
        removed.push(relPath);
      }
    }

    const notesRelsDir = path.join(notesDir, "_rels");
    if (pyfs.exists(notesRelsDir)) {
      for (const filePath of pyfs.glob(notesRelsDir, "*.rels")) {
        const notesFile = path.join(
          notesDir,
          path.basename(filePath).replace(".rels", "")
        );
        if (!pyfs.exists(notesFile)) {
          fs.unlinkSync(filePath);
          removed.push(pyfs.relativeTo(unpackedDir, filePath));
        }
      }
    }
  }

  return removed;
}

function updateContentTypes(unpackedDir, removedFiles) {
  const ctPath = path.join(unpackedDir, "[Content_Types].xml");
  if (!pyfs.exists(ctPath)) {
    return;
  }

  const dom = minidom.parseFile(ctPath);
  let changed = false;

  const removedSet = new Set(removedFiles);
  for (const override of getElementsByTagNameArray(dom, "Override")) {
    const partName = override.getAttribute("PartName").replace(/^\/+/, "");
    if (removedSet.has(partName)) {
      if (override.parentNode) {
        override.parentNode.removeChild(override);
        changed = true;
      }
    }
  }

  if (changed) {
    fs.writeFileSync(ctPath, minidom.toxml(dom, "utf-8"));
  }
}

function cleanUnusedFiles(unpackedDir) {
  const allRemoved = [];

  const slidesRemoved = removeOrphanedSlides(unpackedDir);
  allRemoved.push(...slidesRemoved);

  const trashRemoved = removeTrashDirectory(unpackedDir);
  allRemoved.push(...trashRemoved);

  while (true) {
    const removedRels = removeOrphanedRelsFiles(unpackedDir);
    const referenced = getReferencedFiles(unpackedDir);
    const removedFiles = removeOrphanedFiles(unpackedDir, referenced);

    const totalRemoved = removedRels.concat(removedFiles);
    if (totalRemoved.length === 0) {
      break;
    }

    allRemoved.push(...totalRemoved);
  }

  if (allRemoved.length) {
    updateContentTypes(unpackedDir, allRemoved);
  }

  return allRemoved;
}

function main() {
  const argv = process.argv.slice(2);
  if (argv.length !== 1) {
    process.stderr.write("Usage: node clean.js <unpacked_dir>\n");
    process.stderr.write("Example: node clean.js unpacked/\n");
    process.exit(1);
  }

  const unpackedDir = argv[0];

  if (!pyfs.exists(unpackedDir)) {
    process.stderr.write(`Error: ${unpackedDir} not found\n`);
    process.exit(1);
  }

  const removed = cleanUnusedFiles(unpackedDir);

  if (removed.length) {
    console.log(`Removed ${removed.length} unreferenced files:`);
    for (const f of removed) {
      console.log(`  ${f}`);
    }
  } else {
    console.log("No unreferenced files found");
  }
}

if (require.main === module) {
  main();
}

module.exports = { cleanUnusedFiles };
