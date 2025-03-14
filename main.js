const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Hyperlink, ImageRun } = require("docx");

// Create the main window
function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  win.loadFile("index.html");
}

// Parse Markdown and convert to DOCX elements
function parseMarkdown(markdown) {
  const lines = markdown.split("\n");
  const docElements = [];
  let inCodeBlock = false;
  let inBlockquote = false;
  let inOrderedList = false;
  let inUnorderedList = false;

  for (const line of lines) {
    if (!line.trim()) {
      docElements.push(new Paragraph(""));
      continue;
    }

    // Handle code blocks (```)
    if (line.startsWith("```")) {
      inCodeBlock = !inCodeBlock;
      continue;
    }

    if (inCodeBlock) {
      docElements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: line,
              font: "Courier New",
            }),
          ],
        })
      );
      continue;
    }

    // Handle blockquotes (>)
    if (line.startsWith("> ")) {
      docElements.push(
        new Paragraph({
          text: line.replace(/^>\s*/, ""),
          style: "Blockquote",
        })
      );
      continue;
    }

    // Handle horizontal rules (--- or ***)
    if (/^[-*]{3,}$/.test(line)) {
      docElements.push(
        new Paragraph({
          text: "",
          thematicBreak: true,
        })
      );
      continue;
    }

    // Handle headers (#, ##, ###, etc.)
    const headerMatch = line.match(/^(#+)\s(.*)/);
    if (headerMatch) {
      const level = headerMatch[1].length;
      const text = headerMatch[2];
      docElements.push(
        new Paragraph({
          text: text,
          heading: HeadingLevel[`HEADING_${Math.min(level, 6)}`],
        })
      );
      continue;
    }

    // Handle unordered lists (*, -, +)
    if (/^[\*\-+]\s/.test(line)) {
      docElements.push(
        new Paragraph({
          text: line.replace(/^[\*\-+]\s/, ""),
          bullet: { level: 0 },
        })
      );
      continue;
    }

    // Handle ordered lists (1., 2., etc.)
    if (/^\d+\.\s/.test(line)) {
      docElements.push(
        new Paragraph({
          text: line.replace(/^\d+\.\s/, ""),
          numbering: { level: 0, reference: "ordered-list" },
        })
      );
      continue;
    }

    // Handle bold (**text**)
    if (/\*\*(.*?)\*\*/.test(line)) {
      const text = line.replace(/\*\*(.*?)\*\*/g, "$1");
      docElements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: text,
              bold: true,
            }),
          ],
        })
      );
      continue;
    }

    // Handle italic (*text* or _text_)
    if (/[*_](.*?)[*_]/.test(line)) {
      const text = line.replace(/[*_](.*?)[*_]/g, "$1");
      docElements.push(
        new Paragraph({
          children: [
            new TextRun({
              text: text,
              italic: true,
            }),
          ],
        })
      );
      continue;
    }

    // Handle links ([text](url))
    if (/\[.*?\]\(.*?\)/.test(line)) {
      const linkMatch = line.match(/\[(.*?)\]\((.*?)\)/);
      if (linkMatch) {
        const text = linkMatch[1];
        const url = linkMatch[2];
        docElements.push(
          new Paragraph({
            children: [
              new Hyperlink({
                children: [new TextRun(text)],
                link: url,
              }),
            ],
          })
        );
        continue;
      }
    }

    // Handle images (![alt](url))
    if (/!\[.*?\]\(.*?\)/.test(line)) {
      const imageMatch = line.match(/!\[(.*?)\]\((.*?)\)/);
      if (imageMatch) {
        const altText = imageMatch[1];
        const imageUrl = imageMatch[2];
        docElements.push(
          new Paragraph({
            children: [
              new ImageRun({
                data: fs.readFileSync(imageUrl),
                transformation: { width: 200, height: 200 }, // Adjust size as needed
              }),
            ],
          })
        );
        continue;
      }
    }

    // Default paragraph
    docElements.push(new Paragraph(line));
  }

  return docElements;
}

// Handle conversion in main process
ipcMain.handle("convert-md-to-docx", async (event, markdown) => {
  try {
    const doc = new Document({
      sections: [{
        children: parseMarkdown(markdown),
      }],
      numbering: {
        config: [
          {
            reference: "ordered-list",
            levels: [{ level: 0, format: "decimal", text: "%1.", alignment: "left" }],
          },
        ],
      },
    });

    // Save to exported-documents folder
    const exportDir = path.join(__dirname, "exported-documents");
    if (!fs.existsSync(exportDir)) fs.mkdirSync(exportDir);
    
    const filename = `converted-${Date.now()}.docx`;
    const filePath = path.join(exportDir, filename);
    
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);
    
    return { success: true, filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Start the app
app.whenReady().then(createWindow);
app.on("window-all-closed", () => app.quit());