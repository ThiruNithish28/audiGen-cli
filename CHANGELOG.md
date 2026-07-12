# Changelog

## [v0.1.1] - 2026-07-12

### Added
- LLM fallback mechanism to handle API rate limit error (429) from free-tier providers

### Fixed
- Fixed date format issue (start date , end date ) in the generated Excel file
- Fix TestCase B4 Ticket ID issue

---

## [v0.1.0] - 2025-04-26

### Added
- Interactive dual-mode CLI (flags or questionary prompts)
- Config registry with `auditgen config setup`
- Generates Impact Analysis, Test Cases, Code Checklist from BRD
- Windows EXE via GitHub Actions
- Input validation with friendly error messages
- Path traversal protection on ticket ID
- Ctrl+C handling across all prompts