# Testing

## Prerequisites

Install development dependencies:

```bash
pip install -e ".[dev]"
```

## Running Tests

Run the full test suite:

```bash
pytest
```

Run with verbose output:

```bash
pytest -v
```

Run a specific test file:

```bash
pytest tests/test_powerpoint.py
pytest tests/test_excel.py
pytest tests/test_word.py
pytest tests/test_cli.py
pytest tests/test_utils.py
pytest tests/test_file_ops.py
```

## Test Structure

```
tests/
├── conftest.py          # Shared fixtures
├── test_cli.py          # CLI entry point tests
├── test_utils.py        # Utility function tests (_parse_color, _resolve_data_path, etc.)
├── test_powerpoint.py   # PowerPoint creation/editing tests
├── test_excel.py        # Excel creation/editing tests
├── test_word.py         # Word document creation/editing tests
└── test_file_ops.py     # File operation tests (get_file, download_media)
```

## Key Fixtures (conftest.py)

| Fixture | Description |
|---|---|
| `tmp_data_dir` | Creates a temporary directory and patches `DATA_PATH` to use it. Restores the original value after the test. |
| `pptx_file` | Creates a test PowerPoint file and returns its path. Depends on `tmp_data_dir`. |
| `xlsx_file` | Creates a test Excel file and returns its path. Depends on `tmp_data_dir`. |
| `docx_file` | Creates a test Word document and returns its path. Depends on `tmp_data_dir`. |
| `sample_image` | Creates a small red 100x100 PNG image and returns its path. Depends on `tmp_data_dir`. |

## Data Path in Tests

The `tmp_data_dir` fixture patches `mcp_server.DATA_PATH` to point to a pytest-managed temporary directory (`tmp_path`). This ensures:

- Tests do not write to the real default data path (`/tmp/onit-office-<pid>`)
- Each test gets an isolated directory
- Cleanup is handled automatically by pytest

When no `--data-path` is supplied on the CLI, the server defaults to a process-specific temporary directory (`<system temp dir>/onit-office-<pid>`), which is automatically cleaned up on server exit. See the [README](../README.md#data-directory) for details.
