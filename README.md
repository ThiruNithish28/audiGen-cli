# AudiGen CLI

Audit document generator CLI tool — generates Impact Analysis, Test Cases, 
and Code Review Checklist from a BRD document using AI.

## Requirements
- Windows 10/11
- Gemini API key ([get one free here](https://aistudio.google.com/))

## Installation
1. Download `auditgen.exe` from [Releases](../../releases)
2. Place it in a folder e.g. `C:\Tools\auditgen\`
3. Add that folder to your Windows PATH
4. Open a new terminal and run `auditgen --help`

## First Time Setup
```cmd
auditgen config setup
```
Select all fields and enter your details when prompted.

## Usage
```cmd
# Interactive mode — prompts for everything
auditgen generate

# Direct mode — pass everything as flags
auditgen generate "path\to\brd.docx" TKT-001 -s 20-04-2025 -e 30-04-2025

# View your config
auditgen config show
```

## Output
Running `generate` produces three Excel files in your output folder:
- `TKT-001-Impact Analysis Template.xlsx`
- `TKT-001-Test Cases.xlsx`  
- `TKT-001-Code Checklist.xlsx`

## Built With
- Python 3.12
- Click — CLI framework
- Google Gemini — test case generation
- openpyxl — Excel generation
- Rich + Questionary — terminal UI