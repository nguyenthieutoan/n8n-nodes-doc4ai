# Doc4AI — Agent Guide

> This document helps AI Agents understand and use this node when building n8n workflows.

## Overview

| Property | Value |
|----------|-------|
| **Package** | `n8n-nodes-doc4ai` |
| **Node Type** | Regular Node |
| **Connection Type** | Input/Output: `Main` |
| **Credential** | None (free API) |
| **n8n displayName** | `Doc4AI` |

## What This Node Does

Extracts text from documents (PDF, DOC, DOCX, XLS, XLSX, CSV) and converts text back to documents. Useful for document processing pipelines, especially when feeding content into AI agents or storing AI-generated content as documents.

## Credentials Setup

No credentials required — Doc4AI provides a free document conversion API.

## Connection Types

| Direction | Type | Description |
|-----------|------|-------------|
| Input | `Main` | Receives items with binary data (files) or text data |
| Output | `Main` | Returns extracted text or generated document binary |

## How to Use in Workflows

### Pattern 1: Extract Text from PDF for AI Processing
```
Read Binary File → Doc4AI (Extract) → AI Agent
```

### Pattern 2: Convert AI Output to Document
```
AI Agent → Doc4AI (Create) → Send Email (attachment)
```

## Gotchas & Known Issues

- **Bundled Dependencies**: Uses `esbuild` instead of `tsc` to bundle runtime deps (`pdf-parse`, `mammoth`, `xlsx`) into dist. This is an exception to the zero-dependency rule.
- **Binary Data**: Input files must be in n8n binary data format. Use "Read Binary File" or HTTP Request nodes to prepare files.
- **File Size**: Large files may take longer to process. Consider timeout settings.
