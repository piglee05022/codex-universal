# codex-universal

`codex-universal` is a reference implementation of the base Docker image available in [OpenAI Codex](http://platform.openai.com/docs/codex).

This repository is intended to help developers cutomize environments in Codex, by providing a similar image that can be pulled and run locally. This is not an identical environment but should help for debugging and development.

For more details on environment setup, see [OpenAI Codex](http://platform.openai.com/docs/codex).

## Usage

The Docker image is available at:

```
docker pull ghcr.io/openai/codex-universal:latest
```

The below script shows how can you approximate the `setup` environment in Codex:

```
docker run --rm -it \
    # See below for environment variable options.
    -e CODEX_ENV_PYTHON_VERSION=3.12 \
    -e CODEX_ENV_NODE_VERSION=20 \
    -e CODEX_ENV_RUST_VERSION=1.87.0 \
    -e CODEX_ENV_GO_VERSION=1.23.8 \
    -e CODEX_ENV_SWIFT_VERSION=6.1 \
    # Mount the current directory similar to how it would get cloned in.
    -v $(pwd):/workspace/$(basename $(pwd)) -w /workspace/$(basename $(pwd)) \
    ghcr.io/openai/codex-universal:latest
```

`codex-universal` includes setup scripts that look for `CODEX_ENV_*` environment variables and configures the language version accordingly.

### Configuring language runtimes

The following environment variables can be set to configure runtime installation. Note that a limited subset of versions are supported (indicated in the table below):

| Environment variable       | Description                | Supported versions                               | Additional packages                                                  |
| -------------------------- | -------------------------- | ------------------------------------------------ | -------------------------------------------------------------------- |
| `CODEX_ENV_PYTHON_VERSION` | Python version to install  | `3.10`, `3.11.12`, `3.12`, `3.13`                | `pyenv`, `poetry`, `uv`, `ruff`, `black`, `mypy`, `pyright`, `isort` |
| `CODEX_ENV_NODE_VERSION`   | Node.js version to install | `18`, `20`, `22`                                 | `corepack`, `yarn`, `pnpm`, `npm`                                    |
| `CODEX_ENV_RUST_VERSION`   | Rust version to install    | `1.83.0`, `1.84.1`, `1.85.1`, `1.86.0`, `1.87.0` |                                                                      |
| `CODEX_ENV_GO_VERSION`     | Go version to install      | `1.22.12`, `1.23.8`, `1.24.3`                    |                                                                      |
| `CODEX_ENV_SWIFT_VERSION`  | Swift version to install   | `5.10`, `6.1`                                    |                                                                      |

## What's included

In addition to the packages specified in the table above, the following packages are also installed:

- `ruby`: 3.2.3
- `bun`: 1.2.10
- `java`: 21
- `bazelisk` / `bazel`

See [Dockerfile](Dockerfile) for the full details of installed packages.

## Legal document generator module

This repository includes a simple example module located in `legal_module/` that demonstrates how to generate Traditional Chinese legal filings using the `python-docx` package. The module exposes a `create_filing` function that accepts case information and outputs a formatted Word document.  When `docx2pdf` is available the file can also be exported to PDF. Generated documents use **2.5&nbsp;cm margins**, **1.5&nbsp;line spacing**, and include an optional **附件** section for supporting materials.

### Features

- Produce common filings such as **起訴狀**, **答辯狀**, **上訴狀**, **強制執行聲請狀**, and **補充說明狀**.
- Output in Traditional Chinese using **標楷體 16pt** font.
- Automatic paragraph numbering and formatting.
- Evidence and attachment lists can be integrated directly into the document.
- Export to **DOCX** and, when Microsoft Word is available, to **PDF** via `docx2pdf`.

Example usage:

```bash
python3 -m legal_module.example
```

### Interactive menu

The module also provides a simple CLI that lists available actions. Run:

```bash
python3 -m legal_module.cli
```

This will display a text menu where you can generate the sample filing or exit the program.

This requires the optional dependencies **python-docx** and **docx2pdf** for PDF export. Install them with:

```bash
pip install python-docx docx2pdf
```

### Viewing generated documents

The produced `.docx` and `.pdf` files can be opened in common office suites:

- [Microsoft Office 365](https://www.microsoft.com/en-us/microsoft-365)
- [WPS Office](https://www.wps.com)
- [LibreOffice](https://www.libreoffice.org)

Simply open the generated file with your preferred program and the formatting will be preserved.
