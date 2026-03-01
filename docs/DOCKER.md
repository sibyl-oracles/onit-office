# Docker

## Build

```bash
docker build -t onit-office .
```

## Run

```bash
docker run -d -p 18203:18203 -v onit-office-data:/root/.onit-office/data --name onit-office onit-office
```

The MCP server SSE endpoint will be available at `http://localhost:18203/sse`.

By default, created files are stored in a temporary directory inside the container (`/tmp/onit-office-<pid>`) and are **automatically cleaned up** when the server exits. To persist files, use `--data-path` with a volume mount:

```bash
docker run -d -p 18203:18203 -v onit-office-data:/data --name onit-office onit-office --data-path /data
```

The named volume (`onit-office-data`) ensures files persist across container restarts.

### Bind mount to a local directory

To access created files directly on your host:

```bash
docker run -d -p 18203:18203 -v $(pwd)/output:/data --name onit-office onit-office --data-path /data
```

Files will appear in `./output/`.

### Custom port

```bash
docker run -d -p 9000:9000 --name onit-office onit-office --port 9000
```

## Stop

```bash
docker stop onit-office
docker rm onit-office
```

## MCP client configuration

Point your MCP client to the SSE endpoint:

```json
{
  "mcpServers": {
    "onit-office": {
      "url": "http://localhost:18203/sse"
    }
  }
}
```
