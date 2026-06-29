"use strict";
/*
 * Create thumbnail grids from PowerPoint presentation slides.
 * (Port of thumbnail.py)
 *
 * Creates a grid layout of slide thumbnails for quick visual analysis.
 * Labels each thumbnail with its XML filename (e.g., slide1.xml).
 * Hidden slides are shown with a placeholder pattern.
 *
 * Usage:
 *   node thumbnail.js input.pptx [output_prefix] [--cols N]
 *
 * NOTE: The control flow, slide selection, grid geometry, file naming and
 * chunking are a faithful port of the Python original. Pixel-level output
 * cannot be byte-identical to Pillow because the underlying raster engine
 * (node-canvas/Cairo) differs from Pillow for font shaping, image
 * down-sampling and JPEG encoding.
 */

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");
const AdmZip = require("adm-zip");
const { createCanvas, loadImage } = require("canvas");
const minidom = require("./_pylib/minidom");
const pyfs = require("./_pylib/pyfs");
const { getSofficeEnv } = require("./office/soffice");

const THUMBNAIL_WIDTH = 300;
const CONVERSION_DPI = 100;
const MAX_COLS = 6;
const DEFAULT_COLS = 3;
const JPEG_QUALITY = 95;
const GRID_PADDING = 20;
const BORDER_WIDTH = 2;
const FONT_SIZE_RATIO = 0.1;
const LABEL_PADDING_RATIO = 0.4;

function stem(p) {
  const b = path.basename(p);
  const e = path.extname(b);
  return e ? b.slice(0, -e.length) : b;
}

async function main() {
  const argv = process.argv.slice(2);
  const positional = [];
  let colsArg = DEFAULT_COLS;
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--cols" || a.startsWith("--cols=")) {
      const v = a.startsWith("--cols=") ? a.slice("--cols=".length) : argv[++i];
      colsArg = parseInt(v, 10);
    } else {
      positional.push(a);
    }
  }
  const input = positional[0];
  const outputPrefix = positional[1] !== undefined ? positional[1] : "thumbnails";

  let cols = Math.min(colsArg, MAX_COLS);
  if (colsArg > MAX_COLS) {
    console.log(`Warning: Columns limited to ${MAX_COLS}`);
  }

  const inputPath = input;
  if (!pyfs.exists(inputPath) || path.extname(inputPath).toLowerCase() !== ".pptx") {
    process.stderr.write(`Error: Invalid PowerPoint file: ${input}\n`);
    process.exit(1);
  }

  const outputPath = `${outputPrefix}.jpg`;

  let tempDir;
  try {
    const slideInfo = getSlideInfo(inputPath);

    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "thumb-"));
    const visibleImages = convertToImages(inputPath, tempDir);

    if (visibleImages.length === 0 && !slideInfo.some((s) => s.hidden)) {
      process.stderr.write("Error: No slides found\n");
      process.exit(1);
    }

    const slides = await buildSlideList(slideInfo, visibleImages, tempDir);

    const gridFiles = await createGrids(slides, cols, THUMBNAIL_WIDTH, outputPath);

    console.log(`Created ${gridFiles.length} grid(s):`);
    for (const gridFile of gridFiles) {
      console.log(`  ${gridFile}`);
    }
  } catch (e) {
    process.stderr.write(`Error: ${e.message}\n`);
    process.exit(1);
  } finally {
    if (tempDir) {
      try {
        fs.rmSync(tempDir, { recursive: true, force: true });
      } catch (e) {}
    }
  }
}

function getSlideInfo(pptxPath) {
  const zip = new AdmZip(pptxPath);

  const relsContent = zip.readAsText("ppt/_rels/presentation.xml.rels", "utf8");
  const relsDom = minidom.parseString(relsContent);

  const ridToSlide = {};
  const rels = relsDom.getElementsByTagName("Relationship");
  for (let i = 0; i < rels.length; i++) {
    const rel = rels[i];
    const rid = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    const relType = rel.getAttribute("Type");
    if (relType.indexOf("slide") !== -1 && target.startsWith("slides/")) {
      ridToSlide[rid] = target.replace("slides/", "");
    }
  }

  const presContent = zip.readAsText("ppt/presentation.xml", "utf8");
  const presDom = minidom.parseString(presContent);

  const slides = [];
  const sldIds = presDom.getElementsByTagName("p:sldId");
  for (let i = 0; i < sldIds.length; i++) {
    const sldId = sldIds[i];
    const rid = sldId.getAttribute("r:id");
    if (Object.prototype.hasOwnProperty.call(ridToSlide, rid)) {
      const hidden = sldId.getAttribute("show") === "0";
      slides.push({ name: ridToSlide[rid], hidden });
    }
  }

  return slides;
}

async function buildSlideList(slideInfo, visibleImages, tempDir) {
  let placeholderSize;
  if (visibleImages.length) {
    const img = await loadImage(visibleImages[0]);
    placeholderSize = [img.width, img.height];
  } else {
    placeholderSize = [1920, 1080];
  }

  const slides = [];
  let visibleIdx = 0;

  for (const info of slideInfo) {
    if (info.hidden) {
      const placeholderPath = path.join(tempDir, `hidden-${info.name}.jpg`);
      const placeholderCanvas = createHiddenPlaceholder(placeholderSize);
      fs.writeFileSync(placeholderPath, placeholderCanvas.toBuffer("image/jpeg", { quality: 1 }));
      slides.push([placeholderPath, `${info.name} (hidden)`]);
    } else {
      if (visibleIdx < visibleImages.length) {
        slides.push([visibleImages[visibleIdx], info.name]);
        visibleIdx += 1;
      }
    }
  }

  return slides;
}

function createHiddenPlaceholder(size) {
  const [w, h] = size;
  const canvas = createCanvas(w, h);
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#F0F0F0";
  ctx.fillRect(0, 0, w, h);

  const lineWidth = Math.max(5, Math.floor(Math.min(w, h) / 100));
  ctx.strokeStyle = "#CCCCCC";
  ctx.lineWidth = lineWidth;

  ctx.beginPath();
  ctx.moveTo(0, 0);
  ctx.lineTo(w, h);
  ctx.stroke();

  ctx.beginPath();
  ctx.moveTo(w, 0);
  ctx.lineTo(0, h);
  ctx.stroke();

  return canvas;
}

function convertToImages(pptxPath, tempDir) {
  const pdfPath = path.join(tempDir, `${stem(pptxPath)}.pdf`);

  let result = spawnSync(
    "soffice",
    ["--headless", "--convert-to", "pdf", "--outdir", String(tempDir), String(pptxPath)],
    { encoding: "utf-8", env: getSofficeEnv() }
  );
  if (result.status !== 0 || !pyfs.exists(pdfPath)) {
    throw new Error("PDF conversion failed");
  }

  result = spawnSync(
    "pdftoppm",
    ["-jpeg", "-r", String(CONVERSION_DPI), String(pdfPath), String(path.join(tempDir, "slide"))],
    { encoding: "utf-8" }
  );
  if (result.status !== 0) {
    throw new Error("Image conversion failed");
  }

  const rx = /^slide-.*\.jpg$/;
  return fs
    .readdirSync(tempDir)
    .filter((n) => rx.test(n))
    .map((n) => path.join(tempDir, n))
    .sort();
}

async function createGrids(slides, cols, width, outputPath) {
  const maxPerGrid = cols * (cols + 1);
  const gridFiles = [];

  let chunkIdx = 0;
  for (let startIdx = 0; startIdx < slides.length; startIdx += maxPerGrid) {
    const endIdx = Math.min(startIdx + maxPerGrid, slides.length);
    const chunkSlides = slides.slice(startIdx, endIdx);

    const grid = await createGrid(chunkSlides, cols, width);

    let gridFilename;
    if (slides.length <= maxPerGrid) {
      gridFilename = outputPath;
    } else {
      const s = stem(outputPath);
      const suffix = path.extname(outputPath);
      const parent = path.dirname(outputPath);
      gridFilename = path.join(parent, `${s}-${chunkIdx + 1}${suffix}`);
    }

    const parentDir = path.dirname(gridFilename);
    fs.mkdirSync(parentDir, { recursive: true });
    fs.writeFileSync(gridFilename, grid.toBuffer("image/jpeg", { quality: JPEG_QUALITY / 100 }));
    gridFiles.push(gridFilename);

    chunkIdx += 1;
  }

  return gridFiles;
}

async function createGrid(slides, cols, width) {
  const fontSize = Math.trunc(width * FONT_SIZE_RATIO);
  const labelPadding = Math.trunc(fontSize * LABEL_PADDING_RATIO);

  const first = await loadImage(slides[0][0]);
  const aspect = first.height / first.width;
  const height = Math.trunc(width * aspect);

  const rows = Math.floor((slides.length + cols - 1) / cols);
  const gridW = cols * width + (cols + 1) * GRID_PADDING;
  const gridH = rows * (height + fontSize + labelPadding * 2) + (rows + 1) * GRID_PADDING;

  const grid = createCanvas(gridW, gridH);
  const ctx = grid.getContext("2d");
  ctx.fillStyle = "white";
  ctx.fillRect(0, 0, gridW, gridH);

  ctx.font = `${fontSize}px sans-serif`;
  ctx.textBaseline = "top";

  for (let i = 0; i < slides.length; i++) {
    const [imgPath, slideName] = slides[i];
    const row = Math.floor(i / cols);
    const col = i % cols;
    const x = col * width + (col + 1) * GRID_PADDING;
    const yBase = row * (height + fontSize + labelPadding * 2) + (row + 1) * GRID_PADDING;

    const label = slideName;
    const textW = Math.round(ctx.measureText(label).width);
    ctx.fillStyle = "black";
    ctx.fillText(label, x + Math.floor((width - textW) / 2), yBase + labelPadding);

    const yThumbnail = yBase + labelPadding + fontSize + labelPadding;

    const img = await loadImage(imgPath);
    const [w, h] = thumbnailSize(img.width, img.height, width, height);
    const tx = x + Math.floor((width - w) / 2);
    const ty = yThumbnail + Math.floor((height - h) / 2);
    ctx.drawImage(img, tx, ty, w, h);

    if (BORDER_WIDTH > 0) {
      ctx.strokeStyle = "gray";
      ctx.lineWidth = BORDER_WIDTH;
      // Mirror PIL rectangle coords: inclusive bottom-right corner.
      const x0 = tx - BORDER_WIDTH;
      const y0 = ty - BORDER_WIDTH;
      const x1 = tx + w + BORDER_WIDTH - 1;
      const y1 = ty + h + BORDER_WIDTH - 1;
      ctx.strokeRect(x0, y0, x1 - x0, y1 - y0);
    }
  }

  return grid;
}

// Replicate PIL Image.thumbnail((boxW, boxH)): shrink-only, preserve aspect.
function thumbnailSize(w, h, boxW, boxH) {
  const factor = Math.min(boxW / w, boxH / h);
  if (factor >= 1) return [w, h];
  const nw = Math.max(Math.round(w * factor), 1);
  const nh = Math.max(Math.round(h * factor), 1);
  return [nw, nh];
}

if (require.main === module) {
  main();
}

module.exports = {
  getSlideInfo,
  buildSlideList,
  createHiddenPlaceholder,
  convertToImages,
  createGrids,
  createGrid,
};
