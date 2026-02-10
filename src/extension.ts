import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";

const VIDEO_EXTENSIONS = new Set([
  ".mp4",
  ".mov",
  ".wmv",
  ".avi",
  ".mkv",
  ".m4v",
  ".mpeg",
  ".mpg"
]);

function getNonConflictingPath(dir: string, fileName: string): string {
  const ext = path.extname(fileName);
  const base = path.basename(fileName, ext);

  let candidate = path.join(dir, fileName);
  let counter = 1;

  while (fs.existsSync(candidate)) {
    candidate = path.join(dir, `${base} (${counter})${ext}`);
    counter++;
  }

  return candidate;
}

export function activate(context: vscode.ExtensionContext) {
  const disposable = vscode.commands.registerCommand(
    "pptxVideoExtractor.extractVideos",
    async (uri: vscode.Uri) => {
      if (!uri?.fsPath || path.extname(uri.fsPath).toLowerCase() !== ".pptx") {
        vscode.window.showErrorMessage("Please select a PPTX file.");
        return;
      }

      try {
        const pptxPath = uri.fsPath;

        const folderPick = await vscode.window.showOpenDialog({
          canSelectFiles: false,
          canSelectFolders: true,
          canSelectMany: false,
          openLabel: "Select output folder"
        });

        if (!folderPick || folderPick.length === 0) {
          vscode.window.showInformationMessage("Extraction cancelled.");
          return;
        }

        const outputDir = folderPick[0].fsPath;

        const data = fs.readFileSync(pptxPath);
        const zip = await JSZip.loadAsync(data);

        const mediaFolder = "ppt/media/";
        const mediaFiles = Object.keys(zip.files).filter(
          (filePath) => filePath.startsWith(mediaFolder) && !zip.files[filePath].dir
        );

        if (mediaFiles.length === 0) {
          vscode.window.showInformationMessage("No media files found in this PPTX.");
          return;
        }

        let extractedCount = 0;

        for (const filePath of mediaFiles) {
          const ext = path.extname(filePath).toLowerCase();
          if (!VIDEO_EXTENSIONS.has(ext)) continue;

          const fileName = path.basename(filePath);
          const destPath = getNonConflictingPath(outputDir, fileName);

          const content = await zip.files[filePath].async("nodebuffer");
          fs.writeFileSync(destPath, content);

          extractedCount++;
        }

        if (extractedCount === 0) {
          vscode.window.showInformationMessage("No video files found in this PPTX.");
        } else {
          vscode.window.showInformationMessage(
            `Extracted ${extractedCount} video file(s) to: ${outputDir}"
          );
        }
      } catch (err: any) {
        vscode.window.showErrorMessage(`Failed to extract videos: ${err.message ?? err}`);
      }
    }
  );

  context.subscriptions.push(disposable);
}

export function deactivate() {}