# PDF Page Merge App

A responsive web app for merging PDF files with page-level drag-and-drop sorting.

## Features

- Upload multiple PDF files at once.
- Rearrange individual pages with drag-and-drop.
- Remove any page before export.
- Merge and download directly in the browser.
- Mobile-friendly layout and controls.
- Static build that is easy to host on Netlify.

## Local Development

1. Install dependencies:

```bash
npm install
```

2. Start the dev server:

```bash
npm run dev
```

3. Build for production:

```bash
npm run build
```

## Deploy To Netlify

This project includes a `netlify.toml` file, so Netlify can auto-detect build settings.

If you connect this repo in Netlify:

- Build command: `npm run build`
- Publish directory: `dist`

After deploy, the app is fully static and runs in-browser.
