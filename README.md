# Doc Viewer (PWA)

Open **PowerPoint, PDF, Word, Excel, CSV, text and image** files full-screen on your phone. Installs with its own icon and works offline.

## Deploy on GitHub (with build.yml)

1. Create a new **public** repository on GitHub.
2. Upload everything from this ZIP (extract it first). Keep the same names:
   - `index.html, style.css, app.js, pptx.js, ppt-legacy.js, readers.js, jszip.min.js, sw.js`
   - `manifest.webmanifest, icon-192.png, icon-512.png, icon-maskable-512.png`
   - `package.json, build.mjs`
   - **`.github/workflows/build.yml`**  (the build workflow)
3. **Phone tip:** the GitHub web upload hides folders that start with a dot. Instead, use **Add file → Create new file**, type the name `.github/workflows/build.yml` and paste the contents of `build.yml` (included at the top of the ZIP). Commit.
4. Go to **Settings → Pages → Build and deployment → Source: "GitHub Actions"**.
5. Open the **Actions** tab. The workflow "Build and deploy" runs by itself (about 1–2 minutes). When it is green, your app is live at
   `https://YOUR-USERNAME.github.io/REPO-NAME/`
6. Open that link in Chrome → menu (⋮) → **Install app**. The icon appears on your home screen.

What the workflow does: installs `pdfjs-dist` with npm, copies the PDF engine into `dist/vendor`, and publishes `dist/` to GitHub Pages. Without this step PDFs will not open (everything else still works).

## Using the app

- **+ Add files**: pick one or many files. **Add a whole folder**: adds every supported file inside it.
- Added files stay in the app. Tap a file once to open it.
- **Slides (.pptx):** tap = next slide, left edge or swipe = previous. Menu (☰, top-left): slider, auto-play, full screen.
- **PDF:** scrolls like a normal document; menu → *Page mode* turns it into tap-to-turn pages. Zoom with the menu or pinch.
- **Word (.docx), Excel (.xlsx), CSV, text, images:** scroll view with zoom.
- File manager → file → **Share → Doc Viewer** also opens files directly (after installing the app).

## Limits

- A web app cannot list all files on a phone by itself, so files or folders must be added once.
- Old **.ppt** files show **text only** (the old binary format has no web-friendly layout). Save as .pptx for the full design.
- Old **.doc / .xls** are not supported. Save as .docx / .xlsx.
- Charts and SmartArt in slides show as a dashed box. Animations and transitions are not played.
- Word files are shown as a reflowing page (good for phones), so page breaks and headers/footers differ from Word.
