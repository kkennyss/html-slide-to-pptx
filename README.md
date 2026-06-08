# HTML Slide to PPTX

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Convert HTML slide decks to editable PowerPoint presentations — with a pixel-perfect fallback.**
> **把 HTML 幻灯片转为可编辑的 PPTX，截图兜底保视觉一致。**

An OpenClaw agent skill that turns HTML slides into **real, editable PPTX files** using a dual-method approach:

- **Primary (editable):** Uses `dom-to-pptx` (browser CDN) — text boxes, shapes, images preserve as native OOXML elements. User opens in PowerPoint and can click to edit.
- **Fallback (pixel-perfect):** Headless Chrome screenshots → `python-pptx` assembly — when layout fidelity matters more than editability.

一个 OpenClaw agent 技能。双路线：主路线生成可编辑 PPTX，兜底路线截图保视觉一致。

## Features

| Feature | Primary (dom-to-pptx) | Fallback (screenshot) |
|---------|-----------------------|-----------------------|
| Editable in PowerPoint | ✅ Shapes, text, images | ❌ Each slide is one image |
| Visual fidelity | ~99% (may show OOXML warning on first open) | ✅ Pixel-perfect |
| Dependencies | None (browser only) | Chrome + python-pptx |
| Setup | Open HTML, click button | `pip install -r requirements.txt` |

## Quick Start

### Method 1: Editable PPTX (dom-to-pptx)

Design HTML slides as `.slide` divs at 1280×720px, include the CDN script:

```html
<script src="https://cdn.jsdelivr.net/npm/dom-to-pptx@1.1.10/dist/dom-to-pptx.bundle.js"></script>
<button onclick="domToPptx.exportToPptx(
  Array.from(document.querySelectorAll('.slide')),
  { fileName:'out.pptx', layout:'LAYOUT_16x9', autoEmbedFonts:false, svgAsVector:false }
)">Download PPTX</button>
```

Then serve the HTML (`python -m http.server` or any server), open in browser, click **Download PPTX**.

### Method 2: Screenshot Fallback

```bash
pip install requests websocket-client python-pptx
python scripts/screenshot_to_pptx.py http://127.0.0.1:8000/demo.html --output output.pptx
```

## Layout Design Principles

### 1. Feel over form
A slide looks "饱满" when its content block is vertically centered and substantial — regardless of layout type. Let content dictate, not formula.

### 2. Padding compensates for content volume
- Heavy content (dense): 30-40px padding-top
- Medium (cards): 50-60px
- Light (single table/paragraph): 60-70px
- Minimal (notes): 15-20px

### 3. Context enriches
Thin slide? Add a lead sentence, key data point, or summary line instead of forcing format changes.

### 4. Cover & section slides are exceptions

## Project Structure

```
html-slide-to-pptx/
├── SKILL.md              # OpenClaw agent skill definition
├── scripts/
│   └── screenshot_to_pptx.py  # Chrome screenshot → PPTX converter
├── examples/
│   └── demo.html         # Example HTML slide deck
├── LICENSE
└── README.md
```

## Requirements (Screenshot Fallback)

- Python 3.10+
- Chrome/Chromium
- pip packages: `requests`, `websocket-client`, `python-pptx`

## License

MIT © [kkennyss](https://github.com/kkennyss)
