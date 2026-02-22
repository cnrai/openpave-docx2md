# openpave-docx2md

📄 Convert DOCX and DOC files to Markdown format.

## Installation

```bash
# From local directory
pave install ~/path/to/openpave-docx2md

# From GitHub
pave install cnrai/openpave-docx2md
```

## Prerequisites

For best results, install **pandoc**:

```bash
# macOS
brew install pandoc

# Ubuntu/Debian
sudo apt install pandoc

# Windows
choco install pandoc
```

The skill will fall back to native XML parsing for .docx files if pandoc is not available, but the output quality may be lower.

**.doc files require pandoc** - there is no native fallback for the legacy format.

## Usage

### Convert a Document

```bash
# Convert and print to stdout
docx2md convert document.docx

# Convert and save to file
docx2md convert document.docx -o output.md

# Shorthand (file as first argument)
docx2md document.docx -o output.md

# Convert without line wrapping (better for editing)
docx2md convert document.docx --nowrap -o output.md

# Extract embedded images
docx2md convert document.docx --extract-media=./images -o output.md

# Use reference-style links
docx2md convert document.docx --reference -o output.md

# Force native parsing (skip pandoc)
docx2md convert document.docx --native -o output.md
```

### Get Document Info

```bash
# Human-readable info
docx2md info document.docx

# JSON output
docx2md info document.docx --json
```

## Commands

| Command | Description |
|---------|-------------|
| `convert <file>` | Convert DOCX/DOC to Markdown |
| `info <file>` | Show document information |
| `help` | Show help message |

## Convert Options

| Option | Description |
|--------|-------------|
| `-o, --output <file>` | Output file (default: stdout) |
| `--wrap <mode>` | Line wrapping: `auto`, `none`, `preserve` |
| `--nowrap` | Disable line wrapping (same as `--wrap=none`) |
| `--standalone` | Include YAML front matter with metadata |
| `--extract-media <dir>` | Extract images to specified directory |
| `--reference` | Use reference-style links instead of inline |
| `--native` | Force native XML parsing (skip pandoc) |
| `--summary` | Show conversion summary |

## Info Output

When using `docx2md info`, you'll see:

- File name and path
- File size
- Modified date
- Document title (if available)
- Author (if available)
- Word count (if available)
- Page count (if available)
- Pandoc availability

## Examples

### Basic Conversion

```bash
# Simple conversion
docx2md convert report.docx -o report.md

# Check the result
cat report.md
```

### Batch Conversion

```bash
# Convert all DOCX files in a directory
for f in *.docx; do
  docx2md convert "$f" -o "${f%.docx}.md"
done
```

### With Images

```bash
# Extract images while converting
docx2md convert manual.docx --extract-media=./manual-images -o manual.md

# Images will be referenced in markdown as:
# ![](./manual-images/image1.png)
```

### Clean Output for Editing

```bash
# No line wrapping, reference links (cleaner for git diffs)
docx2md convert document.docx --nowrap --reference -o document.md
```

## Conversion Quality

### With Pandoc (Recommended)

- Full formatting support (headings, lists, tables, links)
- Proper image extraction
- Metadata preservation
- Support for both .doc and .docx

### Native Fallback (DOCX only)

- Basic paragraph extraction
- Heading detection
- Bold and italic text
- Limited formatting support
- No image extraction

## Troubleshooting

### "pandoc not found"

Install pandoc for your platform:

```bash
# macOS
brew install pandoc

# Linux
sudo apt install pandoc
```

### ".doc files require pandoc"

The legacy .doc format cannot be parsed natively. Install pandoc to convert .doc files.

### "Invalid DOCX file"

The file may be corrupted or not a valid DOCX. Try opening it in Word first to verify.

## License

MIT
