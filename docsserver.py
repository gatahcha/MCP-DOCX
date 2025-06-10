# docsserver.py
from fastmcp import FastMCP  # ✅ Correct import for FastMCP 2.0
from docseditor import DocsEditor  # Import your existing class
import os
from typing import Optional

# Create MCP server
mcp = FastMCP("DOCX Document Creator")

# Global document editor instance
current_editor: Optional[DocsEditor] = None

@mcp.tool
def create_new_document() -> str:
    """Create a new DOCX document with default formatting"""
    global current_editor
    current_editor = DocsEditor()
    return "✅ New document created with Times New Roman formatting and standard margins."

@mcp.tool
def add_document_title(title: str) -> str:
    """Add a title to the document (22pt, bold, centered, with underline)"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    current_editor.add_title(title)
    return f"✅ Title added: '{title}' (22pt, bold, centered with full-width underline)"

@mcp.tool
def add_paragraph_text(text: str) -> str:
    """Add a paragraph to the document (12pt, justified, Times New Roman)"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    current_editor.add_paragraph(text)
    return f"✅ Paragraph added: '{text[:50]}...'" if len(text) > 50 else f"✅ Paragraph added: '{text}'"

@mcp.tool
def add_section_heading(heading: str) -> str:
    """Add a section heading to the document (18pt, not bold, justified)"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    current_editor.add_heading(heading, level=1)
    return f"✅ Heading added: '{heading}' (18pt, Times New Roman)"

@mcp.tool
def add_source_citation(dates: str, full_source: str, shorten_source: str) -> str:
    """Add a source citation with clickable link (format: dates, shortened_source)"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    current_editor.add_source(dates, full_source, shorten_source)
    return f"✅ Source added: ({dates}, {shorten_source}) - clickable link to {full_source}"

@mcp.tool
def add_document_footer() -> str:
    """Add footer with document title and copyright notice"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    current_editor.add_footer()
    return "✅ Footer added with document title and copyright notice (centered)"

@mcp.tool
def save_document(filename: str) -> str:
    """Save the document as a DOCX file in the document/ directory"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    # Ensure .docx extension
    if not filename.endswith('.docx'):
        filename += '.docx'
    
    # Create document directory if it doesn't exist
    doc_dir = "document"
    if not os.path.exists(doc_dir):
        os.makedirs(doc_dir)
        print(f"Created directory: {doc_dir}")
    
    # Create full file path
    full_path = os.path.join(doc_dir, filename)
    
    # Save the document
    current_editor.save(full_path)
    
    # Check if file was created
    if os.path.exists(full_path):
        file_size = os.path.getsize(full_path)
        return f"✅ Document saved successfully: {full_path} ({file_size} bytes)"
    else:
        return f"❌ Failed to save document: {full_path}"

@mcp.tool
def set_custom_margins(top: float = 1.0, bottom: float = 1.0, left: float = 0.75, right: float = 0.75) -> str:
    """Set custom document margins (in inches)"""
    global current_editor
    if current_editor is None:
        return "❌ No document created. Use create_new_document() first."
    
    current_editor.set_margins(top, bottom, left, right)
    return f"✅ Margins set - Top: {top}\", Bottom: {bottom}\", Left: {left}\", Right: {right}\""


@mcp.tool
def get_document_status() -> str:
    """Get current document status and information"""
    global current_editor
    if current_editor is None:
        return "❌ No document currently loaded. Use create_new_document() to start."
    
    title = current_editor._title if current_editor._title else "No title set"
    return f"✅ Document active - Title: '{title}' | Ready for content addition"

@mcp.tool
def get_formatting_guide() -> str:
    """Get complete formatting guide for DOCX documents"""
    return """


🎯 DOCX Document Formatting Guide

📄 Document Specifications:
• Title: 22pt Times New Roman, bold, centered, full-width underline
• Headings: 18pt Times New Roman, not bold, justified
• Paragraphs: 12pt Times New Roman, justified
• Sources: 10pt Times New Roman, right-aligned, format: (date, link)
• Footer: Centered, title + copyright
• Margins: Top/Bottom 1", Left/Right 0.75"
• Save Location: document/ directory

🔧 Available Tools:
1. create_new_document() - Start new document
2. add_document_title(title) - Add main title
3. add_section_heading(heading) - Add section headers
4. add_paragraph_text(text) - Add body content
5. add_source_citation(dates, full_source, short_source) - Add citations
6. add_document_footer() - Add footer
7. save_document(filename) - Save as DOCX in document/ folder
8. create_housing_report_example() - Quick example

💡 Workflow:
1. Create document → 2. Add title → 3. Add content → 4. Add footer → 5. Save
"""

if __name__ == "__main__":
    mcp.run()  # This handles the stdio transport correctly