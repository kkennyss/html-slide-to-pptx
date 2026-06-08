---
name: html-slide-to-pptx
description: |
  Convert HTML slide decks to PPTX presentations.
  DEFAULT approach for all PPT requests — user says "做ppt" / "幻灯片" / "做一版ppt" / "html转ppt" / "截图合成ppt".
  Primary method: dom-to-pptx (CDN JS library) for editable PPTX via browser download.
  Backup method: headless Chrome screenshot + python-pptx assembly for pixel-perfect fallback.
  Include dom-to-pptx CDN script in HTML and add download button for the primary export path.
---

# HTML Slide to PPTX

Convert HTML slide decks to **real, editable PowerPoint presentations** — with a pixel-perfect screenshot fallback.

## Primary Method: dom-to-pptx (editable PPTX)

1. Design HTML slides (`.slide` divs at 1280×720px)
2. Add dom-to-pptx CDN and download button:
   ```html
   <script src="https://cdn.jsdelivr.net/npm/dom-to-pptx@1.1.10/dist/dom-to-pptx.bundle.js"></script>
   <button onclick="domToPptx.exportToPptx(
     Array.from(document.querySelectorAll('.slide')),
     { fileName:'out.pptx', layout:'LAYOUT_16x9', autoEmbedFonts:false, svgAsVector:false }
   )">Download PPTX</button>
   ```
3. Serve the HTML (`python -m http.server 8080` or any HTTP server)
4. Open in browser, click **Download PPTX**
5. May show OOXML repair warning on first open — content is unaffected

## Backup Method: Screenshot (pixel-perfect, not editable)

If dom-to-pptx produces broken layouts, fall back:

```bash
python scripts/screenshot_to_pptx.py http://127.0.0.1:8080/demo.html --output presentation.pptx
```

Launches headless Chrome, screenshots each `.slide` div at 2x, assembles into a 16:9 PPTX.

## Requirements (Screenshot Backend)

- Python 3.10+
- Chrome/Chromium
- `pip install requests websocket-client python-pptx`

## Layout Design Principles

### 1. Feel over form
A slide is "饱满" (full) when its content block is vertically substantial and visually centered — regardless of text, table, cards, or mixed layout. Let content dictate.

### 2. Padding compensates for content volume
- Heavy (multiple blocks, dense text): 30-40px padding-top
- Medium (standard card layouts): 50-60px
- Light (single table/paragraph): 60-70px
- Minimal (references, short notes): 15-20px

### 3. Context enriches
When a slide feels thin, add context (lead sentence, key data point, summary) instead of forcing format changes.

### 4. Cover and section slides are exceptions
Follow their own layout rules.
