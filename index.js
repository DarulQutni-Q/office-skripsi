#!/usr/bin/env node

import { platform, arch } from 'os';

console.log(`
┌──────────────────────────────────────────────────────┐
│               SKRIPSI SKILL / OFFICE SKILL           │
│  AI Agent Skill for Thesis DOCX Editing & Audit      │
│                                                        │
│  Platform: ${platform()} (${arch()})                          │
│                                                        │
│  Load SKILL.md in your AI agent to get started.        │
│                                                        │
│  Sub-skills:                                           │
│    skills/docx/       - DOCX editing (main)           │
│    skills/pptx/       - PPTX/PowerPoint               │
│    skills/pdf/        - PDF handling                  │
│    skills/spreadsheet/ - XLSX/Spreadsheets            │
│                                                        │
│  Repo: https://github.com/andypratama3/office-skill    │
└──────────────────────────────────────────────────────┘
`);
