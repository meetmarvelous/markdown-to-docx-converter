const { Document, Packer, Paragraph, TextRun } = require("docx");
const fs = require("fs");
const path = require("path");

module.exports = {
  convertMarkdownToDocx: (markdown) => {
    const paragraphs = markdown.split("\n").map(line => {
      return new Paragraph({
        children: [new TextRun(line)],
      });
    });

    const doc = new Document({
      sections: [
        {
          properties: {},
          children: paragraphs,
        },
      ],
    });

    // Create the "exported-documents" folder if it doesn't exist
    const exportFolder = path.join(__dirname, "exported-documents");
    if (!fs.existsSync(exportFolder)) {
      fs.mkdirSync(exportFolder);
    }

    // Generate a unique filename
    const timestamp = new Date().toISOString().replace(/[:.]/g, "-");
    const filePath = path.join(exportFolder, `converted-${timestamp}.docx`);

    // Save the DOCX file
    return Packer.toBuffer(doc).then(buffer => {
      fs.writeFileSync(filePath, buffer);
      return filePath;
    });
  },
};