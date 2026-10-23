#!/usr/bin/env python3
"""Package a self-contained HTML app and an offline, end-user ZIP."""
from pathlib import Path
import argparse
import re
import zipfile

ROOT = Path(__file__).resolve().parents[1]


def build(output):
    output.mkdir(parents=True, exist_ok=True)
    html = (ROOT / 'index.html').read_text()
    html = html.replace('<link rel="stylesheet" href="styles.css">', '<style>\n' + (ROOT / 'styles.css').read_text() + '\n</style>')
    html = re.sub(r'\s*<script defer src="[^"]+"></script>', '', html)
    html = html.replace('href="index.html"', 'href="New.html"')
    policy = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; connect-src 'none'; object-src 'none'; base-uri 'none'; form-action 'none'"
    html = html.replace('<meta charset="utf-8">', '<meta charset="utf-8">\n  <meta http-equiv="Content-Security-Policy" content="' + policy + '">')
    scripts = []
    for name in ['vendor/xlsx.full.min.js', 'vendor/fflate.js', 'template.js', 'converter.js', 'app.js']:
        # User workbook strings can contain HTML: never let them terminate a script.
        contents = re.sub(r'</script', r'<\\/script', (ROOT / name).read_text(), flags=re.I)
        scripts.append('<script>\n' + contents + '\n</script>')
    html = html.replace('</body>', '\n'.join(scripts) + '\n</body>')
    standalone = output / 'New.html'
    standalone.write_text(html)
    archive_path = output / 'New-Canvas-offline.zip'
    with zipfile.ZipFile(archive_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        archive.write(standalone, 'New-Canvas/New.html')
        archive.write(ROOT / 'LEGGIMI.txt', 'New-Canvas/LEGGIMI.txt')
        for name in ['SHEETJS-LICENSE.txt', 'FFLATE-LICENSE.txt']:
            archive.write(ROOT / 'vendor' / name, 'New-Canvas/licenze/' + name)
    print(archive_path)
    return archive_path


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--output', type=Path, default=ROOT / 'dist')
    build(parser.parse_args().output)
