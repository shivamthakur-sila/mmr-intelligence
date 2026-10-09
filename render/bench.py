"""Render a deck with the true fonts and measure how full each slide is.

    python render/bench.py deck.pptx [out_dir]      # a built deck
    python render/bench.py reference.pdf [out_dir]  # a human deck, for comparison

Writes out_dir/p-NN.png (one per slide), out_dir/sheet-N.png (contact
sheets of 6) and prints, per slide and as medians:

  vfill  fraction of the content area's rows that carry any ink
  ink    fraction of the content area's pixels that are ink

The content area is the slide minus the top quarter (band, title, rule) and
the outer margins, so chrome does not count. Ink is any pixel visibly darker
than paper. Photos count as ink - they are content.

Reference medians measured this way are the bar. These numbers measure
emptiness, not quality: they catch the stranded block and the floating
table, and nothing else. The slides are still judged by looking at them.

Needs Docker (image mmr-render, built from render/Dockerfile) for .pptx, and
Python with Pillow; PyMuPDF (fitz) renders PDFs when present.
"""
import os
import statistics
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw

DPI = 60
INK_LEVEL = 235  # a channel below this is ink
ROW_INK = 0.005  # a row with this share of ink pixels counts as filled


def to_pdf(src: Path, out: Path) -> Path:
    if src.suffix.lower() == ".pdf":
        return src
    out.mkdir(parents=True, exist_ok=True)
    target = out / src.name
    if target.resolve() != src.resolve():
        target.write_bytes(src.read_bytes())
    host = str(out.resolve())
    subprocess.run(
        [
            "docker", "run", "--rm", "-e", "HOME=/tmp", "-v", f"{host}:/data",
            "--entrypoint", "/usr/bin/soffice", "mmr-render",
            "--headless", "--norestore", "--convert-to", "pdf", "--outdir", "/data", f"/data/{src.name}",
        ],
        check=True,
        env={**os.environ, "MSYS_NO_PATHCONV": "1"},
        stdout=subprocess.DEVNULL,
    )
    return out / (src.stem + ".pdf")


def pages(pdf: Path, out: Path) -> list[Path]:
    import fitz  # PyMuPDF

    for old in out.glob("p-*.png"):
        old.unlink()
    doc = fitz.open(pdf)
    files = []
    for i, page in enumerate(doc, 1):
        f = out / f"p-{i:02d}.png"
        page.get_pixmap(dpi=DPI).save(f)
        files.append(f)
    return files


def measure(png: Path) -> tuple[float, float]:
    im = Image.open(png).convert("L")
    w, h = im.size
    box = (int(w * 0.02), int(h * 0.25), int(w * 0.94), int(h * 0.92))
    area = im.crop(box)
    aw, ah = area.size
    px = area.load()
    ink_rows = 0
    ink = 0
    for y in range(ah):
        row = sum(1 for x in range(aw) if px[x, y] < INK_LEVEL)
        ink += row
        if row >= aw * ROW_INK:
            ink_rows += 1
    return ink_rows / ah, ink / (aw * ah)


def sheets(files: list[Path], out: Path) -> None:
    per = 6
    for n in range(0, len(files), per):
        ims = [Image.open(f) for f in files[n : n + per]]
        w = 640
        h = max(int(i.height * w / i.width) for i in ims)
        sheet = Image.new("RGB", (2 * (w + 8), 3 * (h + 22)), "white")
        draw = ImageDraw.Draw(sheet)
        for k, im in enumerate(ims):
            r, c = divmod(k, 2)
            x, y = c * (w + 8), r * (h + 22)
            sheet.paste(im.resize((w, int(im.height * w / im.width))), (x, y + 20))
            draw.text((x + 2, y + 4), files[n + k].stem, fill="red")
        sheet.save(out / f"sheet-{n // per + 1}.png")


def main() -> None:
    src = Path(sys.argv[1])
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else src.parent / (src.stem + "-bench")
    out.mkdir(parents=True, exist_ok=True)
    files = pages(to_pdf(src, out), out)
    rows = [(f.stem, *measure(f)) for f in files]
    for name, vfill, ink in rows:
        print(f"{name}  vfill {vfill:.2f}  ink {ink:.2f}")
    print(
        f"MEDIAN  vfill {statistics.median(r[1] for r in rows):.2f}  "
        f"ink {statistics.median(r[2] for r in rows):.2f}  "
        f"(slides {len(rows)}, under 0.35 vfill: {sum(1 for r in rows if r[1] < 0.35)})"
    )
    sheets(files, out)


if __name__ == "__main__":
    main()
