"""把 PPT 渲染成逐页 PNG 供目检。

用法：
    python render.py <deck.pptx> <输出目录> [--engine libreoffice|powerpoint] [--dpi 80] [--pages 4,18]

libreoffice（默认）：用 LibreOffice 转 PDF，再用 pdftoppm 出图。不占用 PowerPoint，速度快。
powerpoint：用 PowerPoint COM 导出（固定 1600×900，忽略 --dpi），和最终放映效果一致；只关闭自己打开的这份文件。
两种引擎的输出都命名为 p-01.png、p-02.png……（序号是页码）。输出目录里会先删掉旧的 p-*.png，避免和上一轮的图混在一起。
"""
import argparse
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

SOFFICE = Path("C:/Program Files/LibreOffice/program/soffice.exe")

PS_EXPORT = r"""
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject PowerPoint.Application
$pres = $app.Presentations.Open("{deck}", $true, $false, $false)
$pres.Export("{out}", "PNG", 1600, 900)
$pres.Close()
if ($app.Presentations.Count -eq 0) {{ $app.Quit() }}
"""


def collect(src, out, pages):
    """把 src 里的逐页 PNG 按文件名中的页码改名为 out/p-NN.png；pages 给出时只保留这些页。"""
    for f in src.iterdir():
        if f.suffix.lower() != ".png":
            continue
        k = int(re.findall(r"\d+", f.stem)[-1])
        if pages is None or k in pages:
            shutil.move(f, out / f"p-{k:02d}.png")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("out")
    ap.add_argument("--engine", choices=["libreoffice", "powerpoint"], default="libreoffice")
    ap.add_argument("--dpi", type=int, default=80)
    ap.add_argument("--pages", help="只导出这些页，如 4,18")
    args = ap.parse_args()

    deck = Path(args.deck).resolve()
    out = Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("p-*.png"):
        old.unlink()
    pages = {int(p) for p in args.pages.split(",")} if args.pages else None

    # 先输出到临时目录再改名：源文件被 PowerPoint 打开时也能渲染，且不在项目目录留下 PDF
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        img = tmp / "img"
        img.mkdir()
        if args.engine == "powerpoint":
            subprocess.run(["powershell", "-NoProfile", "-Command", PS_EXPORT.format(deck=deck, out=img)], check=True)
        else:
            src = tmp / "deck.pptx"
            shutil.copy(deck, src)
            subprocess.run([str(SOFFICE), "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(src)],
                           check=True, capture_output=True)
            pdf = tmp / "deck.pdf"
            for p in sorted(pages) if pages else [None]:
                rng = ["-f", str(p), "-l", str(p)] if p else []
                subprocess.run(["pdftoppm", "-r", str(args.dpi), "-png", *rng, str(pdf), str(img / "p")], check=True)
        collect(img, out, pages)
    print("\n".join(str(f) for f in sorted(out.glob("p-*.png"))))


if __name__ == "__main__":
    main()
