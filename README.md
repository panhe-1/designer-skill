# Designer Skill

[中文说明](README.zh-CN.md)

Designer Skill is a privacy-safe source repository for a designer-oriented skill.

This repo keeps the reusable skill content public while leaving out personal files and generated project outputs.

## Included

- `SKILL.md`
  A packaged skill entry that describes when and how the skill should be used.
- `Claude-Design-Sys-Prompt_中文翻译(1).txt`
  The bundled Chinese translated source prompt used as the detailed reference.
- `README.md`
  English repository overview.
- `README.zh-CN.md`
  Chinese repository overview.

## Excluded

- generated PPT, PNG, and preview exports
- thesis or other personal source documents
- personalized scripts with hard-coded local paths or user-specific project data

## Why It Looks Lean

That is intentional. The goal is to keep the public repo shareable and reusable without exposing private context from downstream work created with the skill.

## Notes

- `SKILL.md` is the main skill entry point.
- `Claude-Design-Sys-Prompt_中文翻译(1).txt` remains available as the full translated prompt reference.
- Local generated artifacts are ignored through `.gitignore` so they do not get uploaded by accident.
