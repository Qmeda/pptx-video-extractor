# PPTX Video Extractor

A Visual Studio Code extension that extracts embedded video files from PPTX (PowerPoint) documents.

## Features

- Right-click a `.pptx` file in the Explorer to extract videos
- Select an output folder via GUI
- Handles duplicate names by appending ` (1)`, ` (2)`
- Keeps the original media file names from `ppt/media`

## Requirements

- VS Code 1.80 or later
- Node.js (for building locally)

## Usage

1. Open the project folder in VS Code.
2. Run the extension in development mode using `F5`.
3. In the Extension Development Host, right-click a `.pptx` file.
4. Choose **Extract Videos from PPTX**.
5. Select the output folder.

## Build

```bash
npm install
npm run compile
```

## Notes

- PPTX files are ZIP archives; videos are stored under `ppt/media/`.
- Only common video file extensions are extracted.

## License

Specify your license here.

---

日本語版: [README.ja.md](README.ja.md)