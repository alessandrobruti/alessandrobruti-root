# MD -> HTML (brand Impresoft) -> PDF via Playwright/Chromium
import markdown, os, re, subprocess
R = os.path.join(os.path.dirname(__file__), "..")
md = open(os.path.join(R, "Report_LAica.md")).read()
body = markdown.markdown(md, extensions=["tables", "fenced_code"])
css = """
@page { size: A4; margin: 14mm 13mm 16mm 13mm; }
body { font-family: Manrope, Arial, sans-serif; color:#000; font-size: 9.2pt; line-height: 1.42; background:#fff; }
h1 { font-weight: 800; font-size: 22pt; margin: 0 0 4pt; border-left: 10px solid #FDC300; padding-left: 10px; }
h2 { font-weight: 700; font-size: 14pt; margin: 18pt 0 6pt; padding-bottom: 3pt; border-bottom: 3px solid #FDC300; page-break-after: avoid; }
h3 { font-weight: 600; font-size: 11pt; margin: 12pt 0 4pt; page-break-after: avoid; }
table { border-collapse: collapse; width: 100%; margin: 6pt 0 10pt; font-size: 8pt; page-break-inside: auto; }
tr { page-break-inside: avoid; }
th { background: #000; color: #fff; text-align: left; padding: 4pt 5pt; font-weight: 600; }
td { border-bottom: 1px solid #DCDCDC; padding: 3.5pt 5pt; vertical-align: top; }
tr:nth-child(even) td { background: #F4F4F4; }
img { width: 100%; max-width: 100%; margin: 6pt 0; page-break-inside: avoid; }
pre { background: #F4F4F4; border-left: 4px solid #FDC300; padding: 6pt 8pt; font-size: 6.9pt; line-height: 1.3; overflow: hidden; white-space: pre; page-break-inside: avoid; }
code { font-family: 'Roboto Mono', monospace; font-size: 8pt; }
hr { border: 0; border-top: 1px solid #DCDCDC; margin: 12pt 0; }
strong { font-weight: 700; }
em { color: #3A3A3A; }
"""
html = f"<!doctype html><html lang='it'><head><meta charset='utf-8'><title>Laica Dossier RevOps</title><style>{css}</style></head><body>{body}</body></html>"
out = os.path.join(R, "Report_LAica.html"); open(out, "w").write(html)
js = f"""
const {{ chromium }} = require('playwright');
(async () => {{ const b = await chromium.launch(); const p = await b.newPage();
 await p.goto('file://{os.path.abspath(out)}'); await p.waitForTimeout(500);
 await p.pdf({{ path: '{os.path.abspath(os.path.join(R, "Report_LAica.pdf"))}', format: 'A4', printBackground: true,
   displayHeaderFooter: true, headerTemplate: '<span></span>',
   footerTemplate: '<div style="font-family:Manrope,Arial;font-size:7px;width:100%;padding:0 13mm;color:#6E6E6E;display:flex;justify-content:space-between"><span>Impresoft Engage · Laica RevOps assessment · uso interno</span><span class="pageNumber"></span></div>',
   margin: {{ top: '14mm', bottom: '16mm', left: '13mm', right: '13mm' }} }});
 await b.close(); }})();
"""
open("/tmp/claude-0/pdf.js", "w").write(js)
subprocess.run(["node", "/tmp/claude-0/pdf.js"], check=True, env={**os.environ, "NODE_PATH": subprocess.check_output(["npm", "root", "-g"]).decode().strip()})
print("pdf ok")
