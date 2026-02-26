# RapNet Converter

A Next.js web app that converts RapNet CSV exports to a clean formatted XLSX file — entirely in the browser, no server needed.

## Features

- **Scientific notation fix**: `Report #` and `RapNet Price` are stored as plain integers (no E+09 display issue in Excel)
- **Smart media URL selection**:
  - If any `VideoURL 1–5` is present → uses the first video URL
  - If no video URL → uses the first `Image 1–5` that is an actual http URL
  - Falls back to empty if neither exists
- **No backend**: All processing happens client-side in the browser

## Setup & Run

```bash
# 1. Install dependencies
npm install

# 2. Start development server
npm run dev

# 3. Open in browser
# http://localhost:3000
```

## Build for Production

```bash
npm run build
npm start
```

## Column Mapping

| Output Column | Source |
|---|---|
| Stock # | Stock # |
| Report # | Report # (converted to integer) |
| RapNet Price | RapNet Price (converted to integer) |
| Country of Polishing | Country |
| VideoURL | First VideoURL 1–5, else first http Image 1–5 |
| DiamondImage | Always empty |
| Report Comments | Cert comment |
| All others | Direct match from input |
