import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const COLORS = {
  white: "#FFFFFF",
  navy: "#243F8E",
  blue: "#336CFF",
  blueDark: "#24489E",
  blueSoft: "#EEF4FF",
  text: "#202B3C",
  muted: "#657086",
  line: "#DFE5EF",
  photoFill: "#F5F7FA",
  photoLine: "#B8C3D6",
  profit: "#1A7F5A",
  loss: "#C94848",
};

function argValue(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && index + 1 < process.argv.length ? process.argv[index + 1] : fallback;
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

function addText(slide, options) {
  const shape = slide.shapes.add({
    geometry: "textbox",
    name: options.name,
    position: { left: options.left, top: options.top, width: options.width, height: options.height },
    fill: options.fill || "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  shape.text = String(options.text ?? "");
  shape.text.style = {
    typeface: "Microsoft YaHei",
    fontSize: options.fontSize,
    color: options.color || COLORS.text,
    bold: Boolean(options.bold),
    alignment: options.alignment || "left",
    verticalAlignment: options.verticalAlignment || "top",
    lineSpacing: options.lineSpacing || 1.08,
    insets: options.insets || { top: 0, right: 0, bottom: 0, left: 0 },
    autoFit: options.autoFit || "shrinkText",
  };
  return shape;
}

function addRect(slide, options) {
  return slide.shapes.add({
    geometry: options.radius ? "roundRect" : "rect",
    name: options.name,
    position: { left: options.left, top: options.top, width: options.width, height: options.height },
    fill: options.fill,
    line: { style: "solid", fill: options.lineFill || "none", width: options.lineWidth || 0 },
    ...(options.radius ? { borderRadius: options.radius } : {}),
  });
}

function addLine(slide, name, left, top, width, color = COLORS.line, weight = 1) {
  return slide.shapes.add({
    geometry: "line",
    name,
    position: { left, top, width, height: 0 },
    fill: "none",
    line: { style: "solid", fill: color, width: weight },
  });
}

function money(value) {
  const rounded = Math.round(Number(value || 0));
  return rounded < 0
    ? `-¥${Math.abs(rounded).toLocaleString("zh-CN")}`
    : `¥${rounded.toLocaleString("zh-CN")}`;
}

function addRank(slide, rank, left, top) {
  const circle = addRect(slide, {
    name: `rank-${rank}`,
    left,
    top,
    width: 27,
    height: 27,
    fill: rank <= 3 ? COLORS.blue : COLORS.blueSoft,
    lineFill: rank <= 3 ? COLORS.blue : COLORS.blueSoft,
    radius: 14,
  });
  circle.text = String(rank);
  circle.text.style = {
    typeface: "Microsoft YaHei",
    fontSize: 14,
    bold: true,
    color: rank <= 3 ? COLORS.white : COLORS.blueDark,
    alignment: "center",
    verticalAlignment: "middle",
    insets: { top: 0, right: 0, bottom: 0, left: 0 },
  };
}

function addModelItem(slide, item, rank, box, compact, modelImages, wideModels) {
  const modelTextLength = [...String(item.model || "")].length;
  const modelFontSize = compact
    ? modelTextLength > 42 ? 11 : modelTextLength > 30 ? 13 : 16
    : modelTextLength > 32 ? 15 : 18;
  const photoSize = compact ? 58 : 68;
  const isWideProduct = wideModels.has(String(item.model_code));
  const photoWidth = isWideProduct ? (compact ? 120 : 145) : photoSize;
  const photoHeight = isWideProduct ? (compact ? 42 : 48) : photoSize;
  const rankLeft = box.left;
  const photoLeft = box.left + 36;
  const photoTop = box.top + Math.round((box.height - photoHeight) / 2);
  const textLeft = photoLeft + photoWidth + 13;
  const textWidth = box.width - (textLeft - box.left);

  addRank(slide, rank, rankLeft, box.top + 6);
  const imageAsset = modelImages[item.model_code];
  addRect(slide, {
    name: imageAsset ? `photo-frame-${item.model_code}` : `photo-placeholder-${rank}`,
    left: photoLeft,
    top: photoTop,
    width: photoWidth,
    height: photoHeight,
    fill: imageAsset ? COLORS.white : COLORS.photoFill,
    lineFill: COLORS.photoLine,
    lineWidth: 1,
    radius: 8,
  });

  if (imageAsset) {
    slide.images.add({
      blob: imageAsset.blob,
      contentType: "image/png",
      alt: `${item.model_code} 产品缩略图`,
      fit: "contain",
      geometry: "roundRect",
      borderRadius: 7,
      position: { left: photoLeft + 2, top: photoTop + 2, width: photoWidth - 4, height: photoHeight - 4 },
    });
  } else {
    addText(slide, {
      name: `photo-placeholder-label-${rank}`,
      text: "型号图",
      left: photoLeft,
      top: photoTop,
      width: photoWidth,
      height: photoHeight,
      fontSize: 14,
      color: "#9AA6B8",
      alignment: "center",
      verticalAlignment: "middle",
    });
  }

  addText(slide, {
    name: `model-${rank}`,
    text: item.model,
    left: textLeft,
    top: box.top + 4,
    width: textWidth,
    height: compact ? 36 : 38,
    fontSize: modelFontSize,
    bold: true,
    verticalAlignment: "middle",
    lineSpacing: 1.02,
  });
  addText(slide, {
    name: `transaction-${rank}`,
    text: `交易额  ${money(item.transaction)}`,
    left: textLeft,
    top: box.top + (compact ? 43 : 47),
    width: Math.round(textWidth * 0.56),
    height: 24,
    fontSize: 16,
    color: COLORS.muted,
    verticalAlignment: "middle",
  });
  addText(slide, {
    name: `profit-${rank}`,
    text: `净利  ${money(item.net_profit)}`,
    left: textLeft + Math.round(textWidth * 0.56),
    top: box.top + (compact ? 43 : 47),
    width: Math.round(textWidth * 0.44),
    height: 24,
    fontSize: 16,
    color: Number(item.net_profit) < 0 ? COLORS.loss : COLORS.profit,
    bold: true,
    verticalAlignment: "middle",
  });
  addLine(slide, `item-separator-${rank}`, box.left + 36, box.top + box.height - 1, box.width - 36);
}

function addChrome(slide, designer, summary, pageNumber, pageCount, period) {
  slide.background.fill = COLORS.white;
  addRect(slide, { name: "top-bar", left: 0, top: 0, width: 1280, height: 12, fill: COLORS.navy });
  addText(slide, {
    name: "slide-title",
    text: `设计型号表现｜${designer}`,
    left: 58,
    top: 32,
    width: 720,
    height: 50,
    fontSize: 36,
    bold: true,
    verticalAlignment: "middle",
  });
  addText(slide, {
    name: "period-label",
    text: `${period} · 按交易额降序`,
    left: 915,
    top: 43,
    width: 305,
    height: 32,
    fontSize: 17,
    color: COLORS.muted,
    alignment: "right",
    verticalAlignment: "middle",
  });
  addLine(slide, "title-divider", 58, 104, 1164);
  addRect(slide, {
    name: "designer-summary",
    left: 58,
    top: 122,
    width: 1164,
    height: 58,
    fill: COLORS.blueSoft,
    lineFill: COLORS.blueSoft,
    radius: 10,
  });
  addText(slide, {
    name: "summary-name",
    text: `${designer} · ${summary.model_count}个分红型号`,
    left: 78,
    top: 132,
    width: 380,
    height: 37,
    fontSize: 22,
    color: COLORS.blueDark,
    bold: true,
    verticalAlignment: "middle",
  });
  addText(slide, {
    name: "summary-transaction",
    text: `总交易额  ${money(summary.transaction_total)}`,
    left: 500,
    top: 132,
    width: 310,
    height: 37,
    fontSize: 20,
    bold: true,
    verticalAlignment: "middle",
  });
  addText(slide, {
    name: "summary-profit",
    text: `总净利  ${money(summary.net_profit_total)}`,
    left: 820,
    top: 132,
    width: 260,
    height: 37,
    fontSize: 20,
    color: Number(summary.net_profit_total) < 0 ? COLORS.loss : COLORS.profit,
    bold: true,
    verticalAlignment: "middle",
  });
  addText(slide, {
    name: "page-count",
    text: `${pageNumber} / ${pageCount}`,
    left: 1110,
    top: 132,
    width: 90,
    height: 37,
    fontSize: 17,
    color: COLORS.blueDark,
    alignment: "right",
    verticalAlignment: "middle",
  });
}

function addFooter(slide) {
  addText(slide, {
    name: "footer",
    text: "排序规则：同一设计师内按交易额从高到低｜未提供图片的型号保留灰色预留位",
    left: 58,
    top: 684,
    width: 1164,
    height: 24,
    fontSize: 16,
    color: COLORS.muted,
    verticalAlignment: "middle",
  });
}

function addSources(slide, data, imageRoot) {
  slide.speakerNotes.textFrame.setText([
    "[Sources]",
    `- 销售数据：${data.metadata.sales_file}（${data.metadata.sales_sheet}）`,
    `- 型号归属：${data.metadata.designer_file}（${data.metadata.designer_sheet}）`,
    `- 型号图片：${imageRoot}`,
    "- 汇总规则：按设计师和基础型号匹配，同型号多孔距合并，按有效销售降序。",
    "[/Sources]",
  ].join("\n"));
}

function chooseLayout(modelCount) {
  return modelCount > 10 ? { pageSize: 15, columns: 3 } : { pageSize: 10, columns: 2 };
}

function buildDesignerSlides(presentation, designer, summary, modelImages, wideModels, period, data, imageRoot) {
  const { pageSize, columns } = chooseLayout(summary.model_count);
  const models = [...summary.models].sort((a, b) => Number(b.transaction || 0) - Number(a.transaction || 0));
  const pages = [];
  for (let index = 0; index < models.length; index += pageSize) pages.push(models.slice(index, index + pageSize));

  pages.forEach((items, pageIndex) => {
    const slide = presentation.slides.add();
    addChrome(slide, designer, summary, pageIndex + 1, pages.length, period);
    const left = 58;
    const top = 198;
    const gridWidth = 1164;
    const gridHeight = 468;
    const columnGap = columns === 3 ? 18 : 30;
    const columnWidth = (gridWidth - columnGap * (columns - 1)) / columns;
    const rows = Math.ceil(pageSize / columns);
    const rowHeight = gridHeight / rows;
    items.forEach((item, itemIndex) => {
      const column = Math.floor(itemIndex / rows);
      const row = itemIndex % rows;
      addModelItem(
        slide,
        item,
        pageIndex * pageSize + itemIndex + 1,
        {
          left: left + column * (columnWidth + columnGap),
          top: top + row * rowHeight,
          width: columnWidth,
          height: rowHeight,
        },
        columns === 3,
        modelImages,
        wideModels,
      );
    });
    addFooter(slide);
    addSources(slide, data, imageRoot);
    slide.__previewName = `${designer}_第${pageIndex + 1}页`;
  });
}

async function loadModelImages(imageRoot, designer) {
  const images = {};
  const designerDir = path.join(imageRoot, designer);
  const thumbnailDir = path.join(designerDir, "thumbnails");
  let fileNames = [];
  try {
    fileNames = await fs.readdir(thumbnailDir);
  } catch {
    return images;
  }
  for (const fileName of fileNames) {
    if (!fileName.toLowerCase().endsWith(".png")) continue;
    const modelCode = path.basename(fileName, path.extname(fileName));
    const bytes = await fs.readFile(path.join(thumbnailDir, fileName));
    images[modelCode] = { blob: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) };
  }
  return images;
}

function safeName(value) {
  return String(value).replace(/[\\/:*?"<>|]/g, "-");
}

export async function appendSalesSlides(presentation, { data, imageRoot, period, cropConfig }) {
  const wideModels = new Set((cropConfig.wide_models || []).map(String));
  const designerOrder = data.designer_order?.length ? data.designer_order : Object.keys(data.designers || {});
  if (!designerOrder.length) throw new Error("型号归属表中没有可生成的设计师数据。 ");

  const startCount = presentation.slides.items.length;
  for (const designer of designerOrder) {
    const summary = data.designers[designer];
    if (!summary?.models?.length) continue;
    const modelImages = await loadModelImages(imageRoot, designer);
    buildDesignerSlides(presentation, designer, summary, modelImages, wideModels, period, data, imageRoot);
  }
  return {
    designerOrder,
    slideCount: presentation.slides.items.length - startCount,
  };
}

async function main() {
  const inputPath = path.resolve(argValue("--input"));
  const outputPath = path.resolve(argValue("--output"));
  const imageRoot = path.resolve(argValue("--image-root"));
  const previewDir = path.resolve(argValue("--preview-dir"));
  const periodArg = argValue("--period");
  const cropConfigPath = path.resolve(argValue("--crop-config"));
  const data = JSON.parse(await fs.readFile(inputPath, "utf-8"));
  const cropConfig = JSON.parse(await fs.readFile(cropConfigPath, "utf-8"));
  const period = periodArg || data.metadata.period || "本月";

  await fs.mkdir(previewDir, { recursive: true });
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const result = await appendSalesSlides(presentation, { data, imageRoot, period, cropConfig });

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}-${safeName(slide.__previewName || index + 1)}`;
    await writeBlob(path.join(previewDir, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(previewDir, `${stem}.layout.json`), await layout.text());
  }
  const inspect = await presentation.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 100000 });
  await fs.writeFile(path.join(previewDir, "inspect.ndjson"), inspect.ndjson);
  const outputInspectPath = `${outputPath}.inspect.ndjson`;
  await fs.rm(outputInspectPath, { force: true });
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(outputPath);
  // artifact-tool 的 inspect 会按 --output 自动创建同名旁车日志；只在 .build/preview 保留检查结果。
  await fs.rm(outputInspectPath, { force: true });
  console.log(`PPT 已生成：${outputPath}`);
  console.log(`设计师 ${result.designerOrder.length} 人，幻灯片 ${presentation.slides.items.length} 页`);
  process.exit(0);
}

const isDirectRun = process.argv[1]
  && path.resolve(process.argv[1]) === path.resolve(fileURLToPath(import.meta.url));
if (isDirectRun) {
  main().catch((error) => {
    console.error(error?.stack || error);
    process.exitCode = 1;
  });
}
