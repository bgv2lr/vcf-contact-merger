VCF Contact Merger

Python tool to read, merge and normalize contacts from vCard (.vcf) files so they import cleanly into Outlook and iCloud.

Features
- Email, phone, and address parsing
  - Supports standard `EMAIL:` and `TEL:` lines and iCloud `itemX.*` variants
  - Preserves `TYPE` parameters (e.g., `TEL;TYPE=CELL;TYPE=VOICE:`) for correct Outlook/iCloud mapping
  - De-duplicates while keeping order; mobile numbers are prioritized when ordering
- Birthday normalization
  - Converts many formats to `YYYY-MM-DD`; uses `1900` when year is missing (e.g., `07.12` → `1900-12-07`)
- NOTE extraction and cleanup
  - Promotes phones/emails/addresses/title from NOTE lines into proper fields
  - Removes redundant NOTE lines and combines remaining notes into a single NOTE with embedded newlines so Outlook/iCloud display everything
- Duplicate handling
  - Merges duplicate contacts by normalized name, preserving the richer data
- Optional per-contact export
  - Writes one `.vcf` per contact (helpful for Outlook which often imports only the first card in a multi-card file)

Requirements
- Python 3.8+ (no external dependencies)

Configuration
Edit `vcf_config.json` in the project directory. Example:

{
  "input_files": {
    "source": "contacts_private_v13.vcf",
    "update": "icloud.vcf"
  },
  "output_file": "contacts_merged.vcf",
  "backup_enabled": true,
  "backup_suffix": "_backup",
  "log_level": "DEBUG",
  "split_output": true,
  "split_output_dir": "contacts_split",
  "vcf_version": "3.0",
  "phone_validation": {
    "min_digits": 7,
    "check_duplicates": true,
    "allow_international": true
  },
  "conflict_resolution": {
    "auto_resolve": true,
    "prefer_update_for": ["TEL", "ADR", "ORG", "NOTE"],
    "prefer_source_for": ["N", "FN", "BDAY", "EMAIL"]
  }
}

Notes:
- Set `split_output` to `true` to also generate one `.vcf` per contact in `contacts_split/`.
- `vcf_version` can be set to `"2.1"` if a client is picky; `3.0` works well for iCloud and most Outlook versions.
- The code no longer enforces max limits for phones/emails; all unique entries are written.

Usage
- Windows (PowerShell or Git Bash):
  - `python vcf_merger.py`
- The script will:
  - Read `source` and, if provided, `update`
  - Merge contacts with conflict resolution and NOTE promotion/cleanup
  - Write the combined VCF to `output_file`
  - If `split_output` is true, also write one `.vcf` per contact under `split_output_dir`

Quick Switches
- Split output on/off: set `split_output` to `true` (one `.vcf` per contact) or `false` (single file only).
- Validation report: set `validate_after_write` to `true` to generate `…_validation.txt` after writing.
- What counts as an “issue”: control via `validation_flags` in `vcf_config.json`.
  - By default this repo flags only contacts that have neither phone nor email. You can widen it by setting flags to `true` (e.g., `include_mojibake`).
- Merge audit: set `audit_after_merge` to `true` to output `…_merge_audit.csv/.json` that summarizes what changed per contact.
- Trace specific contacts: add exact display names to `trace_contacts` to log parsed/merged TEL/ADR/EMAIL for just those entries.

Outlook/iCloud Tips
- Outlook often imports only the first contact from a multi-card VCF. Use `split_output: true` and drag all generated `.vcf` files into Outlook’s Contacts folder.
- TYPE parameters (`WORK`, `HOME`, `CELL`) are preserved for phone numbers so Outlook/iCloud put them in the right fields.
- Remaining notes are combined into a single NOTE property with line breaks for reliable display.

Suggested .gitignore
__pycache__/
*.pyc
*.pyo
*.log
*_backup*
contacts_split/
*.vcf
!sample/*.vcf
.DS_Store
Thumbs.db
.env

License
This project is provided as-is without warranty. Add a license here if you intend to publish under a specific license.
