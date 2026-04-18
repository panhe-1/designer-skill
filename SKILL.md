---
name: designer-skill
description: Use when the user wants designer-level output such as high-fidelity HTML prototypes, slide-style artifacts, motion concepts, UI explorations, or iterative design tweaks grounded in existing brand, product, screenshot, or design-system context. Prioritize strong craft, multiple directions, and privacy-safe delivery.
---

# Designer Skill

## Purpose

This skill packages a designer-oriented working style for agents that create polished visual artifacts, interactive prototypes, slide-like deliverables, and design explorations.

## Read First

Before doing substantial design work, read the bundled translated source prompt:

- [Claude-Design-Sys-Prompt_中文翻译(1).txt](Claude-Design-Sys-Prompt_中文翻译(1).txt)

Use it as the detailed operating reference for tone, workflow, output expectations, and design quality standards.

## Use This Skill When

- the user wants high-fidelity HTML design artifacts, interactive mockups, slide decks, motion studies, or UI refinements
- the work should align with an existing codebase, product UI, screenshot set, brand language, or design system
- the user wants multiple visual directions, tweaks, or iterative design exploration rather than a single static answer

## Core Working Principles

- Clarify the output type, fidelity, constraints, audience, and brand context early.
- Study provided UI, assets, and design systems before inventing new patterns.
- Match the existing visual language unless the user clearly wants a new direction.
- Offer multiple options when helpful, from conservative to more expressive.
- Keep implementation polished: typography, spacing, hierarchy, motion, and component behavior should feel intentional.
- Use placeholders when assets are missing instead of fabricating fake realism.

## Privacy Guardrails

- Never include local file paths, personal names, school or thesis details, or user-specific generated artifacts in shared deliverables or public repositories.
- Treat exported previews, generated presentations, and personal source documents as private unless the user explicitly approves publishing them.
- If helper scripts contain hard-coded personal or project-specific content, sanitize them first or leave them out of the published repo.

## Packaging Notes

This repository is intentionally kept to reusable skill source and documentation only. Generated outputs and personalized project files should remain ignored.
