# pptx-review

A CLI tool for **programmatic PowerPoint (.pptx) editing** using Microsoft's [Open XML SDK](https://github.com/dotnet/Open-XML-SDK). Takes a `.pptx` file and a JSON edit manifest, produces a modified presentation with text replacements, shape edits, speaker notes, slide manipulation, and comments — no macros, no compatibility issues.

**Ships as a single ~12MB native binary.** No runtime, no Docker required.

## Why Open XML SDK?

We evaluated three approaches for programmatic presentation editing:

| Approach | Text Editing | Slide Ops | Notes/Comments | Formatting |
|----------|:-:|:-:|:-:|:-:|
| **Open XML SDK (.NET)** | ✅ 100% | ✅ 100% | ✅ 100% | ✅ Preserved |
| python-pptx | ✅ Good | ⚠️ Limited | ⚠️ ~80% | ✅ Preserved |
| LibreOffice CLI | ⚠️ Lossy | ❌ None | ❌ None | ⚠️ Degraded |

Open XML SDK is the gold standard — it's Microsoft's own library for manipulating Office documents. Text replacement handles multi-run spans, formatting is always preserved, and all slide operations work correctly.

## Quick Start

### Option 1: Native Binary (recommended)

```bash
git clone https://github.com/henrybloomingdale/pptx-review.git
cd pptx-review
make install    # Builds + installs to /usr/local/bin
```

Requires [.NET 8 SDK](https://dotnet.microsoft.com/download/dotnet/8.0) for building (`brew install dotnet@8`). The resulting binary is self-contained — no .NET runtime needed to run it.

### Option 2: Docker

```bash
make docker     # Builds Docker image
docker run --rm -v "$(pwd):/work" -w /work pptx-review input.pptx edits.json -o edited.pptx
```

### Usage

```bash
# Basic usage
pptx-review input.pptx edits.json -o edited.pptx

# Pipe JSON from stdin
cat edits.json | pptx-review input.pptx -o edited.pptx

# Custom author name
pptx-review input.pptx edits.json -o edited.pptx --author "Dr. Smith"

# Dry run (validate without modifying)
pptx-review input.pptx edits.json --dry-run

# JSON output for pipelines
pptx-review input.pptx edits.json -o edited.pptx --json

# Read presentation content
pptx-review input.pptx --read --json
```

## JSON Manifest Format

```json
{
  "author": "Reviewer Name",
  "changes": [
    {
      "type": "replace_text",
      "find": "old text",
      "replace": "new text"
    },
    {
      "type": "replace_text",
      "find": "slide-specific text",
      "replace": "new text",
      "slide": 2
    },
    {
      "type": "set_text",
      "slide": 1,
      "shape": "Title 1",
      "text": "New Title"
    },
    {
      "type": "set_notes",
      "slide": 1,
      "text": "Speaker notes for slide 1"
    },
    {
      "type": "delete_slide",
      "slide": 5
    },
    {
      "type": "duplicate_slide",
      "slide": 3
    },
    {
      "type": "reorder_slide",
      "slide": 2,
      "position": 5
    },
    {
      "type": "add_slide",
      "layout": "Blank",
      "position": 3
    }
  ],
  "comments": [
    {
      "slide": 1,
      "text": "This slide needs a better title"
    }
  ]
}
```

### Change Types

| Type | Required Fields | Optional Fields | Description |
|------|----------------|-----------------|-------------|
| `replace_text` | `find`, `replace` | `slide` | Find and replace text. If `slide` specified, only that slide; otherwise all slides. |
| `set_text` | `slide`, `shape`, `text` | | Set text content of a named shape (e.g., "Title 1") |
| `set_notes` | `slide`, `text` | | Set or replace speaker notes for a slide |
| `delete_slide` | `slide` | | Delete slide by number (1-indexed) |
| `duplicate_slide` | `slide` | `position` | Duplicate a slide, optionally place at position |
| `reorder_slide` | `slide`, `position` | | Move slide to a new position |
| `add_slide` | | `layout`, `position` | Add new slide (layout name optional, position optional — defaults to end) |

### Comment Format

Each comment needs:
- `slide` — slide number (1-indexed) to attach the comment to
- `text` — the comment content

## CLI Flags

| Flag | Description |
|------|-------------|
| `-o`, `--output <path>` | Output file path (default: `<input>_edited.pptx`) |
| `--author <name>` | Author name for comments (overrides manifest `author`) |
| `--json` | Output results as JSON (for scripting/pipelines) |
| `--dry-run` | Validate the manifest without modifying the presentation |
| `--read` | Read presentation content (combine with `--json` for structured output) |
| `-v`, `--version` | Show version |
| `-h`, `--help` | Show help |

## Read Mode

Extract presentation content as structured JSON:

```bash
pptx-review input.pptx --read --json
```

Output:
```json
{
  "slides": [
    {
      "number": 1,
      "layout": "Title Slide",
      "shapes": [
        {"name": "Title 1", "type": "textbox", "text": "Presentation Title"},
        {"name": "Subtitle 2", "type": "textbox", "text": "By Author Name"}
      ],
      "notes": "Speaker notes here",
      "comments": []
    }
  ],
  "slide_count": 10
}
```

## Build Targets

```
make              # Build native binary for current platform (~12MB, self-contained)
make install      # Build + install to /usr/local/bin
make all          # Cross-compile for macOS ARM64, macOS x64, Linux x64, Linux ARM64
make docker       # Build Docker image
make test         # Run test (requires TEST_PPTX=path/to/presentation.pptx)
make test-read    # Read a presentation (requires TEST_PPTX=path/to/presentation.pptx)
make clean        # Remove build artifacts
make help         # Show all targets
```

## Exit Codes

- `0` — All changes and comments applied successfully (or read mode completed)
- `1` — One or more edits failed (partial success possible)

## JSON Output Mode

With `--json`, the tool outputs structured results:

```json
{
  "input": "presentation.pptx",
  "output": "presentation_edited.pptx",
  "author": "Dr. Smith",
  "changes_attempted": 5,
  "changes_succeeded": 5,
  "comments_attempted": 2,
  "comments_succeeded": 2,
  "success": true,
  "results": [
    { "index": 0, "type": "comment", "success": true, "message": "Comment added to slide 1" },
    { "index": 0, "type": "replace_text", "success": true, "message": "Replaced 3 occurrence(s) across all slides" }
  ]
}
```

## How It Works

1. Copies the input `.pptx` to the output path
2. Opens the presentation using Open XML SDK (`PresentationDocument`)
3. Adds **comments first** (before changes modify the slide structure)
4. Applies changes (text replacement, shape edits, notes, slide operations)
5. Handles multi-run text matching (text spanning multiple XML runs)
6. Preserves original formatting (RunProperties cloned from source)
7. Saves and reports results

### PowerPoint XML Structure

- **Slides** → `SlidePart` objects linked via `SlideIdList`
- **Shapes** → `Shape` elements in `ShapeTree`, identified by `NonVisualDrawingProperties.Name`
- **Text** → `TextBody` → `Paragraph` → `Run` → `Text`
- **Notes** → `NotesSlidePart` with body placeholder
- **Comments** → `SlideCommentsPart` + `CommentAuthorsPart`
- **Layouts** → `SlideLayoutPart` linked through `SlideMasterPart`

## Development

```bash
# Build native binary (requires .NET 8 SDK)
make build

# Build and run locally
dotnet run -- input.pptx edits.json -o edited.pptx

# Read a presentation
dotnet run -- input.pptx --read --json

# Cross-compile all platforms
make all
# → build/osx-arm64/pptx-review  (macOS Apple Silicon)
# → build/osx-x64/pptx-review    (macOS Intel)
# → build/linux-x64/pptx-review  (Linux x64)
# → build/linux-arm64/pptx-review (Linux ARM64)
```

## License

MIT — see [LICENSE](LICENSE).

---

*Built by [CinciNeuro](https://github.com/henrybloomingdale) for AI-assisted presentation editing workflows.*
