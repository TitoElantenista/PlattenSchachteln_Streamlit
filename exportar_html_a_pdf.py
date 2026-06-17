#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
HTML_NORMAL = SCRIPT_DIR / "Schachtel Kontrolle.html"

BROWSER_CANDIDATES = (
    "google-chrome-stable",
    "google-chrome",
    "chromium",
    "chromium-browser",
)


def buscar_browser(browser: str | None) -> str:
    if browser:
        browser_path = shutil.which(browser) or browser
        if Path(browser_path).exists() or shutil.which(browser_path):
            return browser_path
        raise SystemExit(f"No encontre el navegador indicado: {browser}")

    for candidate in BROWSER_CANDIDATES:
        browser_path = shutil.which(candidate)
        if browser_path:
            return browser_path

    raise SystemExit(
        "No encontre Chrome/Chromium. Instala google-chrome-stable o usa "
        "--browser /ruta/al/navegador."
    )


def css_pdf(formato: str, orientacion: str, mostrar_detalles: bool) -> str:
    detalles_css = "" if mostrar_detalles else "details{display:none!important}"
    return f"""
<style id="pdf-export-style">
@page{{
  size:{formato} {orientacion};
  margin:8mm 8mm 10mm;
}}
@media print{{
  *{{
    -webkit-print-color-adjust:exact!important;
    print-color-adjust:exact!important;
  }}
  html,body{{
    background:#fff!important;
    color:#162033!important;
  }}
  body{{
    font-family:Segoe UI,Arial,sans-serif!important;
    font-size:10px!important;
    line-height:1.3!important;
  }}
  header{{
    position:static!important;
    display:block!important;
    padding:0 0 5mm!important;
    margin:0 0 5mm!important;
    border-bottom:1px solid #cfd6df!important;
    background:#fff!important;
  }}
  .actions,.sheet-actions,.toolbar input,.hover-tip{{
    display:none!important;
  }}
  h1{{
    font-size:22px!important;
    margin:0 0 1mm!important;
  }}
  h2{{
    font-size:15px!important;
    margin:6mm 0 3mm!important;
    break-after:avoid;
    page-break-after:avoid;
  }}
  h3{{
    font-size:12px!important;
    margin:0!important;
  }}
  .eyebrow{{
    font-size:9px!important;
    letter-spacing:0!important;
  }}
  .wrap{{
    max-width:none!important;
    margin:0!important;
    padding:0!important;
  }}
  .metrics{{
    display:grid!important;
    grid-template-columns:repeat(5,1fr)!important;
    gap:2.5mm!important;
    margin:4mm 0 5mm!important;
  }}
  .metric{{
    padding:2.5mm!important;
    border-radius:2mm!important;
    break-inside:avoid;
    page-break-inside:avoid;
  }}
  .metric span{{
    font-size:8px!important;
  }}
  .metric strong{{
    font-size:15px!important;
  }}
  .compare{{
    display:grid!important;
    grid-template-columns:1.1fr .9fr!important;
    gap:4mm!important;
    margin:3mm 0 5mm!important;
    align-items:start!important;
  }}
  .note-box{{
    padding:3mm!important;
    border-radius:2mm!important;
  }}
  .note-box h2{{
    margin-top:0!important;
  }}
  .note-box p{{
    margin:1.5mm 0 0!important;
  }}
  .table-wrap{{
    overflow:visible!important;
    border-radius:2mm!important;
    background:#fff!important;
  }}
  table{{
    width:100%!important;
    border-collapse:collapse!important;
    font-size:8px!important;
  }}
  thead{{
    display:table-header-group;
  }}
  tr{{
    break-inside:avoid;
    page-break-inside:avoid;
  }}
  th,td{{
    padding:1.2mm 1.4mm!important;
    white-space:normal!important;
    overflow-wrap:anywhere;
    word-break:normal;
  }}
  th{{
    position:static!important;
    background:#233047!important;
    color:#fff!important;
  }}
  .toolbar{{
    display:block!important;
    margin:8mm 0 3mm!important;
  }}
  .sheet-list{{
    display:block!important;
    break-before:page;
    page-break-before:always;
  }}
  .sheet-card{{
    margin:0 0 7mm!important;
    border-radius:2mm!important;
    overflow:visible!important;
    break-before:auto;
    page-break-before:auto;
    break-inside:avoid;
    page-break-inside:avoid;
  }}
  .sheet-head{{
    display:block!important;
    padding:2.5mm 3mm!important;
    background:#fbfcfd!important;
  }}
  .sheet-meta{{
    gap:1.5mm 4mm!important;
    font-size:8px!important;
  }}
  .image-frame{{
    padding:2mm!important;
    overflow:visible!important;
    background:#eef2f6!important;
  }}
  .image-frame svg{{
    display:block!important;
    min-width:0!important;
    width:100%!important;
    max-height:132mm!important;
    height:auto!important;
    background:#fff!important;
    border:1px solid #cfd6df!important;
  }}
  details{{
    padding:0 3mm 3mm!important;
  }}
  summary{{
    padding:2mm 0!important;
  }}
  {detalles_css}
}}
</style>
"""


def html_para_pdf(
    html_original: Path,
    html_temporal: Path,
    formato: str,
    orientacion: str,
    mostrar_detalles: bool,
) -> None:
    html = html_original.read_text(encoding="utf-8")
    if mostrar_detalles:
        html = re.sub(r"<details(?![^>]*\bopen\b)([^>]*)>", r"<details open\1>", html)

    estilo = css_pdf(formato, orientacion, mostrar_detalles)
    if "</head>" in html:
        html = html.replace("</head>", f"{estilo}\n</head>", 1)
    else:
        html = f"{estilo}\n{html}"

    html_temporal.write_text(html, encoding="utf-8")


def exportar_pdf(
    browser: str,
    html_original: Path,
    pdf_salida: Path,
    formato: str,
    orientacion: str,
    mostrar_detalles: bool,
    timeout: int,
    mantener_temporal: bool,
) -> Path:
    html_original = html_original.resolve()
    pdf_salida = pdf_salida.resolve()
    pdf_salida.parent.mkdir(parents=True, exist_ok=True)

    if not html_original.exists():
        raise SystemExit(f"No existe el HTML de entrada: {html_original}")

    temp_ctx = None
    if mantener_temporal:
        temp_dir = Path(tempfile.mkdtemp(prefix="schachtel_pdf_"))
    else:
        temp_ctx = tempfile.TemporaryDirectory(prefix="schachtel_pdf_")
        temp_dir = Path(temp_ctx.name)

    try:
        html_temporal = temp_dir / html_original.name
        user_data_dir = temp_dir / "chrome-profile"
        html_para_pdf(html_original, html_temporal, formato, orientacion, mostrar_detalles)

        comando = [
            browser,
            "--headless=new",
            "--disable-gpu",
            "--disable-dev-shm-usage",
            "--no-sandbox",
            "--allow-file-access-from-files",
            "--no-pdf-header-footer",
            "--print-to-pdf-no-header",
            f"--user-data-dir={user_data_dir}",
            f"--print-to-pdf={pdf_salida}",
            html_temporal.as_uri(),
        ]
        resultado = subprocess.run(
            comando,
            cwd=html_original.parent,
            text=True,
            capture_output=True,
            timeout=timeout,
            check=False,
        )
        if resultado.returncode != 0:
            sys.stderr.write(resultado.stderr)
            raise SystemExit(f"Chrome fallo al crear el PDF. Codigo: {resultado.returncode}")

        if not pdf_salida.exists() or pdf_salida.stat().st_size == 0:
            raise SystemExit(f"No se genero un PDF valido: {pdf_salida}")

        if mantener_temporal:
            print(f"HTML temporal: {html_temporal}")

        return pdf_salida
    finally:
        if temp_ctx is not None:
            temp_ctx.cleanup()


def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Exporta el HTML de control de la Plattenschachtelung a un PDF legible."
    )
    parser.add_argument(
        "html",
        nargs="?",
        type=Path,
        help="HTML a exportar. Por defecto usa 'Schachtel Kontrolle.html'.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="PDF de salida. Por defecto usa el mismo nombre del HTML con extension .pdf.",
    )
    parser.add_argument(
        "--formato",
        default="A3",
        choices=("A4", "A3", "A2", "Letter", "Legal"),
        help="Tamano de pagina CSS para impresion. Valor por defecto: A3.",
    )
    parser.add_argument(
        "--orientacion",
        default="landscape",
        choices=("landscape", "portrait"),
        help="Orientacion del PDF. Valor por defecto: landscape.",
    )
    parser.add_argument(
        "--detalles",
        action="store_true",
        help="Abre e imprime las tablas de detalle de cada Tafel.",
    )
    parser.add_argument(
        "--browser",
        help="Ruta o nombre del ejecutable Chrome/Chromium.",
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=180,
        help="Tiempo maximo para Chrome, en segundos.",
    )
    parser.add_argument(
        "--mantener-temporal",
        action="store_true",
        help="Conserva el HTML temporal con el CSS de PDF inyectado.",
    )
    return parser


def main() -> None:
    args = construir_parser().parse_args()
    html_entrada = args.html or HTML_NORMAL
    if not html_entrada.is_absolute():
        html_desde_cwd = (Path.cwd() / html_entrada).resolve()
        html_desde_script = (SCRIPT_DIR / html_entrada).resolve()
        html_entrada = html_desde_cwd if html_desde_cwd.exists() else html_desde_script

    pdf_salida = args.output or html_entrada.with_suffix(".pdf")
    browser = buscar_browser(args.browser)
    pdf = exportar_pdf(
        browser=browser,
        html_original=html_entrada,
        pdf_salida=pdf_salida,
        formato=args.formato,
        orientacion=args.orientacion,
        mostrar_detalles=args.detalles,
        timeout=args.timeout,
        mantener_temporal=args.mantener_temporal,
    )
    print(f"PDF generado: {pdf}")


if __name__ == "__main__":
    main()
