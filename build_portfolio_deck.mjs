import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const IMAGE_EXTENSIONS = new Set([".png", ".jpg", ".jpeg", ".webp"]);
const ROLE_NAMES = new Set(["美工", "设计师"]);

function argValue(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && index + 1 < process.argv.length ? process.argv[index + 1] : fallback;
}

function naturalCompare(a, b) {
  return String(a).localeCompare(String(b), "zh-CN", { numeric: true, sensitivity: "base" });
}

async function directoriesOf(folder) {
  return (await fs.readdir(folder, { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .sort((a, b) => naturalCompare(a.name, b.name));
}

async function imagesOf(folder) {
  return (await fs.readdir(folder, { withFileTypes: true }))
    .filter((entry) => entry.isFile() && IMAGE_EXTENSIONS.has(path.extname(entry.name).toLowerCase()))
    .sort((a, b) => naturalCompare(a.name, b.name))
    .map((entry) => path.join(folder, entry.name));
}

export async function discoverPeople(worksRoot) {
  const people = [];
  const rootImages = await imagesOf(worksRoot);
  if (rootImages.length) {
    people.push({ role: "", name: path.basename(worksRoot), images: rootImages });
  }

  for (const firstLevel of await directoriesOf(worksRoot)) {
    const firstPath = path.join(worksRoot, firstLevel.name);
    if (ROLE_NAMES.has(firstLevel.name)) {
      for (const personDir of await directoriesOf(firstPath)) {
        const personPath = path.join(firstPath, personDir.name);
        const images = await imagesOf(personPath);
        if (images.length) people.push({ role: firstLevel.name, name: personDir.name, images });
      }
      continue;
    }

    const directImages = await imagesOf(firstPath);
    if (directImages.length) {
      people.push({ role: "", name: firstLevel.name, images: directImages });
      continue;
    }

    for (const personDir of await directoriesOf(firstPath)) {
      const personPath = path.join(firstPath, personDir.name);
      const images = await imagesOf(personPath);
      if (images.length) people.push({ role: firstLevel.name, name: personDir.name, images });
    }
  }

  const roleRank = (role) => (role === "美工" ? 0 : role === "设计师" ? 1 : 2);
  people.sort((a, b) => roleRank(a.role) - roleRank(b.role) || naturalCompare(a.name, b.name));
  return people;
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".jpg" || ext === ".jpeg") return "image/jpeg";
  if (ext === ".webp") return "image/webp";
  return "image/png";
}

async function imageBytes(filePath) {
  const bytes = await fs.readFile(filePath);
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

function safeName(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, "-");
}

function addText(slide, name, text, position, style) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name,
    position,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = text;
  shape.text.style = style;
  return shape;
}

function layoutFor(count) {
  if (count <= 1) return { cols: 1, rows: 1, maxWidth: 880 };
  if (count === 2) return { cols: 2, rows: 1 };
  if (count <= 4) return { cols: 2, rows: 2 };
  if (count <= 6) return { cols: 3, rows: 2 };
  if (count <= 8) return { cols: 4, rows: 2 };
  return { cols: 5, rows: 2 };
}

function imageFrames(count) {
  const layout = layoutFor(count);
  const area = { left: 64, top: 116, width: 1152, height: 552 };
  const gapX = layout.cols >= 5 ? 14 : 18;
  const gapY = 18;
  let cellWidth = (area.width - gapX * (layout.cols - 1)) / layout.cols;
  const cellHeight = (area.height - gapY * (layout.rows - 1)) / layout.rows;
  if (layout.maxWidth) cellWidth = Math.min(cellWidth, layout.maxWidth);
  const frames = [];
  let placed = 0;
  for (let row = 0; row < layout.rows && placed < count; row += 1) {
    const rowCount = Math.min(layout.cols, count - placed);
    const rowWidth = rowCount * cellWidth + Math.max(0, rowCount - 1) * gapX;
    const rowLeft = area.left + (area.width - rowWidth) / 2;
    for (let col = 0; col < rowCount; col += 1) {
      frames.push({
        left: rowLeft + col * (cellWidth + gapX),
        top: area.top + row * (cellHeight + gapY),
        width: cellWidth,
        height: cellHeight,
      });
      placed += 1;
    }
  }
  return frames;
}

function pageTitle(person) {
  return person.role ? `${person.role}作品展示 · ${person.name}` : `作品展示 · ${person.name}`;
}

async function addPortfolioSlide(presentation, person, pageImages, pageIndex, pageCount, globalPage, period) {
  const slide = presentation.slides.add();
  slide.background.fill = "#FFFFFF";
  addText(
    slide,
    "page-title",
    pageTitle(person),
    { left: 64, top: 24, width: 820, height: 54 },
    { fontSize: 42, bold: true, color: "#0B0F19", fontFamily: "Microsoft YaHei" },
  );
  const pageSuffix = pageCount > 1 ? ` · ${pageIndex + 1}/${pageCount}` : "";
  addText(
    slide,
    "section-marker",
    `${period}${pageSuffix}  /  ${String(globalPage).padStart(2, "0")}`,
    { left: 900, top: 36, width: 316, height: 28 },
    { fontSize: 16, bold: true, color: "#356CFF", alignment: "right", fontFamily: "Microsoft YaHei" },
  );
  slide.shapes.add({
    geometry: "rect",
    name: "title-rule",
    position: { left: 64, top: 88, width: 1152, height: 4 },
    fill: "#356CFF",
    line: { style: "solid", fill: "#356CFF", width: 0 },
  });

  const frames = imageFrames(pageImages.length);
  for (const [index, frame] of frames.entries()) {
    slide.shapes.add({
      geometry: "rect",
      name: `image-frame-${index + 1}`,
      position: frame,
      fill: "#F8FAFC",
      line: { style: "solid", fill: "#DDE4EE", width: 1 },
    });
    const inset = 5;
    slide.images.add({
      blob: await imageBytes(pageImages[index]),
      contentType: contentType(pageImages[index]),
      alt: `${person.name}作品 ${String(pageIndex * 10 + index + 1).padStart(2, "0")}`,
      fit: "contain",
      position: {
        left: frame.left + inset,
        top: frame.top + inset,
        width: frame.width - inset * 2,
        height: frame.height - inset * 2,
      },
    });
  }

  slide.speakerNotes.textFrame.setText(
    `[Sources]\n- 本页作品图片来自用户选择的本地文件夹：${person.role ? `${person.role}/` : ""}${person.name}。`,
  );
  slide.speakerNotes.setVisible(false);
}

export async function appendPortfolioSlides(presentation, { worksRoot, period, startingPage = 1 }) {
  const people = await discoverPeople(worksRoot);
  if (!people.length) {
    throw new Error("作品文件夹中没有找到图片。请先在软件中粘贴美工或设计师作品图片。 ");
  }

  const startCount = presentation.slides.items.length;
  let globalPage = startingPage;
  for (const person of people) {
    const pages = [];
    for (let index = 0; index < person.images.length; index += 10) pages.push(person.images.slice(index, index + 10));
    for (const [pageIndex, pageImages] of pages.entries()) {
      await addPortfolioSlide(presentation, person, pageImages, pageIndex, pages.length, globalPage, period);
      globalPage += 1;
    }
  }
  return {
    people,
    slideCount: presentation.slides.items.length - startCount,
    nextPage: globalPage,
  };
}

async function main() {
  const worksRoot = path.resolve(argValue("--works-root"));
  const outputPath = path.resolve(argValue("--output"));
  const previewDir = path.resolve(argValue("--preview-dir"));
  const period = argValue("--period", "本月");

  await fs.mkdir(previewDir, { recursive: true });
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const result = await appendPortfolioSlides(presentation, { worksRoot, period });

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(path.join(previewDir, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(previewDir, `${stem}.layout.json`), await layout.text(), "utf8");
  }

  const inspect = await presentation.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 80000 });
  await fs.writeFile(path.join(previewDir, "inspect.ndjson"), inspect.ndjson, "utf8");
  const outputInspectPath = `${outputPath}.inspect.ndjson`;
  await fs.rm(outputInspectPath, { force: true });
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(outputPath);
  await fs.rm(outputInspectPath, { force: true });
  console.log(`作品展示 PPT 已生成：${outputPath}`);
  console.log(`人员 ${result.people.length} 人，幻灯片 ${presentation.slides.items.length} 页`);
}

const isDirectRun = process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
if (isDirectRun) {
  main().catch((error) => {
    console.error(error?.stack || error);
    process.exitCode = 1;
  });
}
