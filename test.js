// test.js
const { Document, Packer, Paragraph, TextRun } = require("docx");

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [
            new TextRun("Hello World!"),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  require("fs").writeFileSync("test.docx", buffer);
  console.log("DOCX file created successfully!");
});