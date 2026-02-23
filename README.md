# onit-office

Standalone MCP server for Microsoft Office document creation (PowerPoint, Excel, Word). Runs as an SSE server that can be used with any MCP client.

Vendored from [onit](https://github.com/sibyl-oracles/onit).

## Installation

```bash
pip install onit-office
```

## Usage

Start the server (runs in background):
```bash
onit-office
# or
onit-office start
```

Start in foreground (for debugging):
```bash
onit-office start --foreground
```

Custom host/port:
```bash
onit-office start --host 127.0.0.1 --port 8000
```

Check status:
```bash
onit-office status
```

Stop the server:
```bash
onit-office stop
```

## MCP Client Configuration

Once running, connect your MCP client to the SSE endpoint:

```
http://localhost:18203/sse
```

## Tools

11 tools for creating and editing Office documents:

**PowerPoint (7 tools):**
- `create_presentation` - Create a new 16:9 presentation with title slide
- `add_slide` - Add slides with various layouts (text, bullets, images, two-column)
- `add_table_slide` - Add data tables with formatted headers
- `add_images_slide` - Display multiple images in grid layouts
- `style_slide` - Apply background colors and styling
- `get_presentation_info` - Inspect presentation structure
- `download_media` - Download media files from URLs

**Excel (2 tools):**
- `create_excel` - Create Excel files with headers and data
- `add_excel_rows` - Append rows to existing Excel files

**Word (2 tools):**
- `create_document` - Create Word documents with optional headers/logos
- `add_document_content` - Add headings, paragraphs, bullets, images, tables

## Data Directory

Created files are stored in `~/.onit-office/data/` by default. Override with `--data-path`.

## License

Apache 2.0
