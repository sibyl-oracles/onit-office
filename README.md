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

## Using with onit

[onit](https://github.com/sibyl-oracles/onit) can connect to onit-office as an MCP SSE server. Start onit-office first, then launch onit with the `--mcp-sse` flag:

```bash
onit-office start
onit --mcp-sse http://localhost:18203/sse --web
```

This gives onit access to all 18 office document tools through its web interface.

## MCP Client Configuration

Once running, connect any MCP client to the SSE endpoint:

```
http://localhost:18203/sse
```

Example MCP client config:

```json
{
  "mcpServers": {
    "onit-office": {
      "url": "http://localhost:18203/sse"
    }
  }
}
```

## Tools

18 tools for creating and editing Office documents:

**PowerPoint (9 tools):**
- `create_presentation` - Create a new 16:9 presentation with title slide
- `add_slide` - Add slides with various layouts (text, bullets, images, two-column)
- `add_table_slide` - Add data tables with formatted headers
- `add_images_slide` - Display multiple images in grid layouts
- `style_slide` - Apply background colors and styling
- `get_presentation_info` - Inspect presentation structure
- `read_presentation` - Read full text content from all slides
- `modify_presentation` - Edit text in shapes or delete slides
- `download_media` - Download media files from URLs

**Excel (4 tools):**
- `create_excel` - Create Excel files with headers and data
- `add_excel_rows` - Append rows to existing Excel files
- `read_excel` - Read cell contents from existing Excel files
- `modify_excel_cells` - Edit specific cells in existing Excel files

**Word (4 tools):**
- `create_document` - Create Word documents with optional headers/logos
- `add_document_content` - Add headings, paragraphs, bullets, images, tables
- `read_document` - Read full content from existing Word documents
- `modify_document` - Edit or delete paragraphs in existing Word documents

**General (1 tool):**
- `get_file` - Retrieve a created file as base64-encoded data for download

## Data Directory

Created files are stored in `~/.onit-office/data/` by default. Override with `--data-path`.

## Docker

See [DOCKER.md](DOCKER.md) for Docker build and run instructions.

## License

Apache 2.0
