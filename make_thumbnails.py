from __future__ import annotations

import argparse
import json
from pathlib import Path

from PIL import Image


SUPPORTED = {".png", ".jpg", ".jpeg", ".webp"}


def contain_on_white(image: Image.Image, size: tuple[int, int], padding: int = 10) -> Image.Image:
    canvas = Image.new("RGB", size, "white")
    work = image.convert("RGB")
    work.thumbnail((size[0] - padding * 2, size[1] - padding * 2), Image.Resampling.LANCZOS)
    canvas.paste(work, ((size[0] - work.width) // 2, (size[1] - work.height) // 2))
    return canvas


def find_source(designer_dir: Path, model: str) -> Path | None:
    for suffix in SUPPORTED:
        candidate = designer_dir / f"{model}{suffix}"
        if candidate.exists():
            return candidate
    model_lower = model.lower()
    for candidate in designer_dir.iterdir():
        if candidate.is_file() and candidate.suffix.lower() in SUPPORTED and candidate.stem.lower() == model_lower:
            return candidate
    return None


def main() -> None:
    parser = argparse.ArgumentParser(description="为设计师型号交易净利 PPT 生成产品缩略图")
    parser.add_argument("--image-root", required=True)
    parser.add_argument("--config", required=True)
    parser.add_argument("--refresh", action="store_true")
    args = parser.parse_args()

    image_root = Path(args.image_root).resolve()
    config = json.loads(Path(args.config).read_text(encoding="utf-8"))
    wide_models = {str(value) for value in config.get("wide_models", [])}
    generated = 0
    skipped = 0

    if not image_root.exists():
        raise FileNotFoundError(f"图片根目录不存在：{image_root}")

    for designer_dir in sorted(path for path in image_root.iterdir() if path.is_dir()):
        if designer_dir.name == "thumbnails":
            continue
        output_dir = designer_dir / "thumbnails"
        output_dir.mkdir(parents=True, exist_ok=True)
        crop_map = config.get(designer_dir.name, {})
        source_models = {path.stem for path in designer_dir.iterdir() if path.is_file() and path.suffix.lower() in SUPPORTED}
        source_models.update(str(model) for model in crop_map)
        for model in sorted(source_models):
            output = output_dir / f"{model}.png"
            if output.exists() and not args.refresh:
                skipped += 1
                continue
            source = find_source(designer_dir, model)
            if source is None:
                continue
            with Image.open(source) as image:
                crop = image.crop(tuple(crop_map[model])) if model in crop_map else image.copy()
                size = (480, 160) if model in wide_models else (240, 240)
                contain_on_white(crop, size).save(output, optimize=True)
                generated += 1

    print(f"缩略图：新生成 {generated}，沿用 {skipped}")


if __name__ == "__main__":
    main()
