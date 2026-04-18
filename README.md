# Designer Skill

Designer Skill is a small workspace for design-oriented prompt assets, presentation generation scripts, and visual reference exports.

This repository currently focuses on two related tracks:

- preserving a Chinese translation draft of a Claude design system prompt
- generating thesis defense presentation decks with Python and `python-pptx`

## Repository Contents

- `Claude-Design-Sys-Prompt_中文翻译(1).txt`
  Chinese translation draft of the design-system-oriented prompt source material.
- `build_defense_ppt.py`
  Base version of the thesis defense deck generator.
- `build_defense_ppt_school_template.py`
  Enhanced school-template version of the thesis defense deck generator.
- `template_media/`
  Image assets used by the presentation scripts.
- `template_export/`
  Preview images for the template-style presentation output.
- `defense_export/`
  Preview images for the base defense deck output.
- `defense_v2_export/`
  Preview images for the enhanced defense deck output.

## Requirements

- Python 3.11 or later
- `python-pptx`

Install dependencies with:

```bash
pip install python-pptx
```

## Quick Start

Run the base presentation generator:

```bash
python build_defense_ppt.py
```

Run the school-template presentation generator:

```bash
python build_defense_ppt_school_template.py
```

## Project Structure

This repository is intentionally lightweight and file-based.

- Prompt material is stored as source text for reference and iteration.
- Presentation scripts generate `.pptx` outputs from structured layout code.
- Export folders keep rendered slide previews for quick review without opening PowerPoint.

## Notes

- Some input and output paths in the Python scripts are currently tailored to the original local machine setup. Review them before running on another computer.
- The repository includes exported preview images so the current visual direction can be inspected directly on GitHub.
- If this project evolves into a reusable skill package, the next recommended step is to separate prompt assets, runtime scripts, and generated outputs into clearer subdirectories.

## Status

This repository is an actively organized working snapshot rather than a finalized packaged product.
