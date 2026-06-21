# n8n-nodes-doc4ai

Developed and maintained by **[Jay Nguyen (Nguyễn Thiệu Toàn)](https://nguyenthieutoan.com)**.

🛡️ **[Verified n8n Creator](https://n8n.io/creators/nguyenthieutoan)** | 💼 CEO/Founder of **[GenStaff](https://genstaff.net)**

**Connect with me:**  
[![LinkedIn](https://img.shields.io/badge/LinkedIn-0077B5?style=flat&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/nguyenthieutoan) [![Facebook](https://img.shields.io/badge/Facebook-1877F2?style=flat&logo=facebook&logoColor=white)](https://www.facebook.com/nguyenthieutoan) [![Website](https://img.shields.io/badge/Website-nguyenthieutoan.com-brightgreen?style=flat)](https://nguyenthieutoan.com) [![Email](https://img.shields.io/badge/Email-me%40nguyenthieutoan.com-blue?style=flat)](mailto:me@nguyenthieutoan.com)

---

**Doc4AI** is a high-performance n8n community node designed to convert diverse document formats into clean, structured Markdown, HTML, or Plain Text optimized for LLMs, RAG pipelines, and AI agents. It also allows you to generate standard files (XLSX, CSV, HTML, TXT) back from text or JSON data.

## Features

### 1. Extract Text from File (Binary to Text)
Extract text from various document formats and convert them into one of the three desired formats: **Markdown**, **HTML**, or **Plain Text**.
*   **Supported Formats**: PDF (all versions), DOC, DOCX, XLS, XLSX, CSV, ODS.
*   **Markdown Preservation**: Converts Word tables, headings, bold/italic structures, and Excel sheets into GitHub Flavored Markdown (GFM) tables and elements, perfect for LLM ingestion.
*   **Advanced Features**:
    *   *PDF Page Range*: Extract only specific pages (e.g., `1-3, 5, 8-10`).
    *   *Excel Sheet Selection*: Parse specific sheets or combine all sheets.
    *   *Header Control*: Use the first row as columns or autogenerate Excel letter columns (A, B, C...).
    *   *File Size Protection*: Restrict processing using a maximum file size limit in MB (defaults to 30 MB, larger files will be rejected).
    *   *Metrics*: Returns file byte size, raw extracted character count, and formatted/converted character count.

### 2. Convert Text to File (Text/JSON to Binary)
Generate binary files directly from text strings or JSON payloads.
*   **Output Formats**: CSV, Excel (XLSX), HTML, Plain Text (TXT).
*   **Spreadsheet Generation**: Auto-converts JSON arrays of objects or 2D arrays into sheet rows.
*   **Customization**: Configure custom Excel sheet names or CSV delimiters.

---

## Why Doc4AI? (Anti-Collision Architecture)

Unlike standard document nodes (e.g., nodes relying on `officeparser` or unbundled loaders), **Doc4AI** completely prevents runtime version collision errors such as:
`PDF processing error: [OfficeParser]: The API version "5.6.205" does not match the Worker version "5.3.31".`

This is achieved using a self-contained, compiled bundling system (using `esbuild`) that encapsulates all parsing engines (`pdf-parse`, `mammoth`, `word-extractor`, `xlsx`) in absolute isolation. This ensures zero dependency pollution and zero runtime collisions with the main n8n core packages.

---

## Installation

### Via n8n UI (Recommended)
1. Go to **Settings > Community nodes > Install**.
2. Enter the package name: `n8n-nodes-doc4ai`.
3. Agree to the terms and click **Install**.
4. Restart your n8n instance if self-hosting.

### Via Command Line
Navigate to your n8n directory (usually `~/.n8n/`) and run:
```bash
npm install n8n-nodes-doc4ai
```
Restart n8n to load the node.

---

## License

This project is licensed under the [MIT License](LICENSE).
