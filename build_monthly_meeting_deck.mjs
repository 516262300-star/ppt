import fs from "node:fs/promises";
import path from "node:path";
import process from "node:process";
import { Presentation, PresentationFile } from "@oai/artifact-tool";
import { appendPortfolioSlides } from "./build_portfolio_deck.mjs";
import { appendSalesSlides } from "./build_deck.mjs";

function argValue(name, fallback = "") {
  const index = process.argv.indexOf(name);
  return index >= 0 && index + 1 < process.argv.length ? process.argv[index + 1] : fallback;
}

async function writeBlob(filePath, blob) {
  await fs.writeFile(filePath, new Uint8Array(await blob.arrayBuffer()));
}

async function main() {
  const worksRoot = path.resolve(argValue("--works-root"));
  const salesInput = path.resolve(argValue("--sales-input"));
  const imageRoot = path.resolve(argValue("--image-root"));
  const cropConfigPath = path.resolve(argValue("--crop-config"));
  const outputPath = path.resolve(argValue("--output"));
  const previewDir = path.resolve(argValue("--preview-dir"));
  const data = JSON.parse(await fs.readFile(salesInput, "utf8"));
  const cropConfig = JSON.parse(await fs.readFile(cropConfigPath, "utf8"));
  const period = argValue("--period") || data.metadata.period || "本月";

  await fs.mkdir(previewDir, { recursive: true });
  await fs.mkdir(path.dirname(outputPath), { recursive: true });

  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const portfolio = await appendPortfolioSlides(presentation, { worksRoot, period, startingPage: 1 });
  const sales = await appendSalesSlides(presentation, { data, imageRoot, period, cropConfig });

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    await writeBlob(path.join(previewDir, `${stem}.png`), await presentation.export({ slide, format: "png", scale: 1 }));
    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(path.join(previewDir, `${stem}.layout.json`), await layout.text(), "utf8");
  }

  const inspect = await presentation.inspect({ kind: "slide,textbox,shape,image,notes", maxChars: 150000 });
  await fs.writeFile(path.join(previewDir, "inspect.ndjson"), inspect.ndjson, "utf8");
  const outputInspectPath = `${outputPath}.inspect.ndjson`;
  await fs.rm(outputInspectPath, { force: true });
  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save(outputPath);
  await fs.rm(outputInspectPath, { force: true });

  console.log(`完整月会 PPT 已生成：${outputPath}`);
  console.log(`作品展示 ${portfolio.slideCount} 页，销售额与利润 ${sales.slideCount} 页，共 ${presentation.slides.items.length} 页`);
}

main().catch((error) => {
  console.error(error?.stack || error);
  process.exitCode = 1;
});
