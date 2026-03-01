# Release Notes

## v0.1.0 — Initial Release

First public release of **onit-office**, a standalone MCP (Model Context Protocol) server for creating and editing Microsoft Office documents. Vendored from [onit](https://github.com/sibyl-oracles/onit).

### Highlights

- SSE-based MCP server compatible with any MCP client
- 18 tools covering PowerPoint, Excel, and Word document workflows
- Runs as a background daemon or in foreground mode
- Docker support included

### PowerPoint Tools (9)

| Tool | Description |
|------|-------------|
| `create_presentation` | Create a new 16:9 widescreen presentation with a title slide |
| `add_slide` | Add slides with 7 layout options: `text`, `bullets`, `image`, `text_image`, `bullets_image`, `two_column`, `blank` |
| `add_table_slide` | Add data tables with auto-styled blue headers |
| `add_images_slide` | Display 2–4 images in `horizontal`, `vertical`, or `grid` (2×2) layouts |
| `style_slide` | Apply background colors by name (e.g. `red`, `blue`) or hex code (`#RRGGBB`) |
| `get_presentation_info` | Retrieve slide count, dimensions, aspect ratio, and slide titles |
| `read_presentation` | Extract full text, tables, and notes from all slides |
| `modify_presentation` | Edit shape text or delete slides by index |
| `download_media` | Download images, audio, video, and PDFs from URLs with content-type detection |

### Excel Tools (4)

| Tool | Description |
|------|-------------|
| `create_excel` | Create workbooks with styled headers, data rows, and auto-adjusted column widths |
| `add_excel_rows` | Append rows to an existing workbook |
| `read_excel` | Read headers and rows (up to 500 by default) from any sheet |
| `modify_excel_cells` | Edit specific cells by reference (e.g. `A1`, `B3`) |

### Word Tools (4)

| Tool | Description |
|------|-------------|
| `create_document` | Create documents with optional title, header, footer, and logo |
| `add_document_content` | Append headings, paragraphs, bullet lists, images, tables, or page breaks |
| `read_document` | Extract paragraphs, tables, headers, and footers |
| `modify_document` | Edit paragraph text/style or delete paragraphs by index |

### General Tools (1)

| Tool | Description |
|------|-------------|
| `get_file` | Retrieve any created file as base64-encoded data with MIME type for client download |

### Server & CLI

- **Default endpoint:** `http://0.0.0.0:18203/sse`
- **Commands:** `onit-office start`, `onit-office stop`, `onit-office status`
- **Options:** `--host`, `--port`, `--data-path`, `--foreground`
- PID-managed background process with automatic stale-PID cleanup
- Logs written to `~/.onit-office/server.log`

### Security

- Data directory created with owner-only permissions (`0o700`)
- Downloaded files written with `0o600` permissions
- All file paths confined to the configured data directory

### LLM Compatibility

All list parameters accept both native Python lists and JSON strings, so LLM callers can pass `'["item1", "item2"]'` directly.

### Docker

```bash
docker build -t onit-office .
docker run -p 18203:18203 onit-office
```

See [DOCKER.md](DOCKER.md) for volume mount and bind mount options.

### Test Suite

161 tests covering all 18 tools, utility functions, CLI, and error handling:

```bash
pip install pytest
pytest tests/ -v
```

### Dependencies

- Python ≥ 3.10
- [FastMCP](https://pypi.org/project/fastmcp/) ≥ 2.0.0
- [python-pptx](https://pypi.org/project/python-pptx/) ≥ 1.0.0
- [openpyxl](https://pypi.org/project/openpyxl/) ≥ 3.1.0
- [python-docx](https://pypi.org/project/python-docx/) ≥ 1.0.0
- [Pillow](https://pypi.org/project/Pillow/) ≥ 10.0.0
- [requests](https://pypi.org/project/requests/) ≥ 2.31.0

### License

Apache 2.0
