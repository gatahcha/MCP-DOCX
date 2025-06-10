# 📄 DOCX Document Creator (FastMCP-Powered)

Welcome to the **DOCX Document Creator** – a FastMCP-powered microserver that allows users to generate well-formatted `.docx` documents programmatically with a consistent style guide.

## 🚀 Features

- 📄 Create a new DOCX document with Times New Roman styling
- 📝 Add a document title with 22pt bold centered text and underline
- 🧾 Insert justified paragraphs with 12pt Times New Roman
- 🧩 Add section headings (18pt, justified, not bold)
- 🔗 Include properly formatted clickable source citations
- 📎 Add footers with copyright and title
- 📁 Save documents to a local `document/` directory
- ✍️ Customize margins and view current document status
- 📘 Retrieve a built-in formatting guide

## 🛠 Usage

This app runs as a FastMCP tool server. After launching `docsserver.py`, you can use the provided tools programmatically via the FastMCP protocol.

### Example Tool Calls

```python
create_new_document()
add_document_title("Economic Trends Report")
add_section_heading("Introduction")
add_paragraph_text("This report covers market trends observed during Q1 2025...")
add_source_citation("June 2025", "https://example.com/report", "Market Report")
add_document_footer()
save_document("Q1_Report.docx")
