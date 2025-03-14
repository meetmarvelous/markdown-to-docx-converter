// Get DOM elements
const markdownInput = document.getElementById("markdown-input");
const convertBtn = document.getElementById("convert-btn");

// Handle button click
convertBtn.addEventListener("click", async () => {
  const markdown = markdownInput.value;
  if (markdown.trim()) {
    try {
      const filePath = await window.electronAPI.convertMarkdownToDocx(markdown);
      alert(`DOCX file saved successfully at: ${filePath}`);
    } catch (err) {
      console.error("Failed to convert Markdown to DOCX:", err);
      alert("Failed to convert Markdown to DOCX.");
    }
  } else {
    alert("Please enter some Markdown text.");
  }
});